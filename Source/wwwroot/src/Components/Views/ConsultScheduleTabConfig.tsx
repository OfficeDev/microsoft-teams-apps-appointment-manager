import * as React from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { Input, Header, Checkbox, CheckboxProps, ComponentEventHandler, InputProps } from '@fluentui/react-northstar';

import { RequestStatus } from "../../Apis/ConsultApi";
import { ConsultScheduleTabSettings } from "../../Models/ConsultScheduleTabSettings";
import { withTranslation, WithTranslation } from "react-i18next";
import { AlertHandler } from "../../Models/AlertHandler";
import { getChannelMappingsForTeam, ChannelMapping } from "../../Apis/ChannelApi";
import { CheckboxItem } from "../../Models/ComponentItems";

// component properties
export interface ConsultScheduleTabConfigProps extends WithTranslation {
    alertHandler: AlertHandler;
}

// component state
export interface ConsultScheduleTabConfigState {
    name: string;
    categories: CheckboxItem<ChannelMapping>[];
    token: string;
    teamsContext: microsoftTeams.Context;
    chkAllStatuses: boolean;
    chkAllCategories: boolean;
    statuses: CheckboxItem<RequestStatus>[];
    existingSettings: ConsultScheduleTabSettings;
}

// ConsultScheduleTabConfig component
class ConsultScheduleTabConfig extends React.Component<ConsultScheduleTabConfigProps, ConsultScheduleTabConfigState> {
    constructor(props: ConsultScheduleTabConfigProps) {
        super(props);
        this.state = {
            name: "",
            categories: [],
            token: "",
            teamsContext: null,
            chkAllCategories: false,
            chkAllStatuses: false,
            statuses: [
                { data: RequestStatus.Unassigned, checked: true },
                { data: RequestStatus.Assigned, checked: true },
                { data: RequestStatus.ReassignRequested, checked: true },
                { data: RequestStatus.Completed, checked: false },
            ],
            existingSettings: null,
        };

        // initialize context
        microsoftTeams.getContext((ctx: microsoftTeams.Context) => {
            this.setState({ teamsContext: ctx });
        });

        // get the settings in case of update
        microsoftTeams.settings.getSettings((settings: microsoftTeams.settings.Settings) => {
            // make sure settings exist...if so, this is an update to the tab
            if (settings.contentUrl && settings.contentUrl.length > 0) {
                // parse the configuration out of the url
                let payload = settings.contentUrl;
                payload = payload.substring(payload.lastIndexOf("/") + 1, payload.length);

                // convert it from a base64 encoded json string to an object
                const json = atob(payload);
                const data: ConsultScheduleTabSettings = JSON.parse(json);

                // process existing status settings (channels will be done when loaded)
                const statuses = this.state.statuses;
                for (let i = 0; i < statuses.length; ++i) {
                    statuses[i].checked = data.statuses.includes(statuses[i].data);
                }

                // save the existingSettings, statuses, and name to state
                this.setState({ existingSettings: data, statuses: statuses, name: settings.suggestedDisplayName });
            }
        });

        // perform the SSO request
        const authTokenRequest = {
            successCallback: this.tokenCallback,
            failureCallback: this.tokenFailedCallback,
        };
        microsoftTeams.authentication.getAuthToken(authTokenRequest);
    }

    // token failure callback for SSO
    tokenFailedCallback = (error: string) => {
        console.error(`SSO failed: ${error}`);
        microsoftTeams.appInitialization.notifyFailure({
            reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
            message: this.props.t('errorAuthFailed'),
        });
    };

    // token callback for SSO
    tokenCallback = (token: string) => {
        // fetch channels
        getChannelMappingsForTeam(token, this.state.teamsContext.teamId).then((channelMappings: ChannelMapping[]) => {
            if (channelMappings.length === 0) {
                this.props.alertHandler(this.props.t('errorNoCategoryFound'), "danger");
            }

            // default checked based on existing settings or current channel
            const categories: CheckboxItem<ChannelMapping>[] = channelMappings.map(mapping => {
                let checked = false;
                if (this.state.existingSettings && this.state.existingSettings.categories.includes(mapping.category)) {
                    checked = true;
                } else if (!this.state.existingSettings && mapping.channelId === this.state.teamsContext.channelId) {
                    checked = true;
                }

                return { data: mapping, checked };
            });

            this.setState({ categories });

            // Tell Teams we are ready to display view
            microsoftTeams.appInitialization.notifySuccess();

            // register handler for saving the configuration
            microsoftTeams.settings.registerOnSaveHandler((saveEvent: microsoftTeams.settings.SaveEvent) => {
                // build payload for configuration
                const payload: ConsultScheduleTabSettings = {
                    categories: this.state.categories.filter(c => c.checked).map(c => c.data.category),
                    statuses: this.state.statuses.filter(s => s.checked).map(s => s.data),
                };
                microsoftTeams.settings.setSettings({
                    contentUrl: `https://${window.location.host}/consult/schedule/${btoa(JSON.stringify(payload))}`,
                    entityId: "schedule",
                    suggestedDisplayName: this.state.name,
                });
                saveEvent.notifySuccess();
            });
        }).catch((err) => {
            // bubble up the error and stop the waiting indicator
            this.props.alertHandler(this.props.t('errorGeneric'), "danger");
            microsoftTeams.appInitialization.notifySuccess();
        });
    };

    // handles input change
    handleInputChange: ComponentEventHandler<InputProps & { value: string; }> = (_evt, data) => {
        this.setState({ name: data.value });
    };

    // handle checkbox changes
    checkChanged = (index: number, t: string, _evt: unknown, ctrl: CheckboxProps) => {
        const collection = (t === "category") ? this.state.categories : this.state.statuses;
        let checkAll = (t === "category") ? this.state.chkAllCategories : this.state.chkAllStatuses;

        if (index === -1) {
            // event raised from the check all checkbox
            checkAll = ctrl.checked;
            if (checkAll) {
                for (let i = 0; i < collection.length; i++) {
                    collection[i].checked = true;
                }
            }
        }
        else {
            collection[index].checked = ctrl.checked;
            if (checkAll && !ctrl.checked) {
                checkAll = false;
            }
        }

        if (t === "category") {
            this.setState({ categories: collection as CheckboxItem<ChannelMapping>[], chkAllCategories: checkAll });
        } else {
            this.setState({ statuses: collection as CheckboxItem<RequestStatus>[], chkAllStatuses: checkAll });
        }
    };

    // renders the component
    render() {
        // enable/disable Teams "save" button based on state
        const isStateValid = this.state.name.length > 2
            && this.state.categories.some(c => c.checked)
            && this.state.statuses.some(s => s.checked);

        microsoftTeams.settings.setValidityState(isStateValid);

        const channelChecks = this.state.categories.map((item, index: number) => (
            <div style={{ width: "100%" }}>
                <Checkbox label={item.data.category} checked={item.checked || this.state.chkAllCategories} onChange={this.checkChanged.bind(this, index, "category")} />
            </div>
        ));

        const statusChecks = this.state.statuses.map((item, index) => (
            <div style={{ width: "100%" }}>
                <Checkbox label={this.statusDisplayNames[item.data]} checked={item.checked || this.state.chkAllStatuses} onChange={this.checkChanged.bind(this, index, "status")} />
            </div>
        ));

        return (
            <div className="page">
                {!this.state.existingSettings &&
                    <Input label={this.props.t('tabNameLabel')} fluid value={this.state.name} onChange={this.handleInputChange.bind(this)} />
                }
                <Header as="h3" content={this.props.t('showRequestsHeader')} />
                <div style={{ width: "100%" }}>
                    <div style={{ width: "50%", float: "left", padding: "5px" }}>
                        <Header as="h4" content={this.props.t('categoriesHeader')} />
                        <div style={{ width: "100%" }}>
                            <Checkbox label={this.props.t('allCheckbox')} onChange={this.checkChanged.bind(this, -1, "category")} checked={this.state.chkAllCategories} />
                        </div>
                        {channelChecks}
                    </div>
                    <div style={{ width: "50%", float: "left", padding: "5px" }}>
                        <Header as="h4" content={this.props.t('statusHeader')} />
                        <div style={{ width: "100%" }}>
                            <Checkbox label={this.props.t('allCheckbox')} onChange={this.checkChanged.bind(this, -1, "status")} checked={this.state.chkAllStatuses} />
                        </div>
                        {statusChecks}
                    </div>
                </div>
            </div>
        );
    }

    private statusDisplayNames: Record<RequestStatus, string> = {
        Unassigned: this.props.t('stateUnassigned'),
        Assigned: this.props.t('stateAssigned'),
        ReassignRequested: this.props.t('stateReassign'),
        Completed: this.props.t('stateCompleted'),
    };
}

export default withTranslation(['consultScheduleTabConfig', 'common'])(ConsultScheduleTabConfig);