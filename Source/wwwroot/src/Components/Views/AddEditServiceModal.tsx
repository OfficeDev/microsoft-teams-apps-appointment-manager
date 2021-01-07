import * as React from "react";
import { RouteComponentProps } from "react-router-dom";
import * as microsoftTeams from "@microsoft/teams-js";
import { Flex, Divider, Input, Text, Dropdown, Button, DropdownProps, ComponentEventHandler, InputProps } from '@fluentui/react-northstar';

import { getChannels, getChannelMappings, postChannelMapping, patchChannelMapping, Channel, ChannelMapping } from '../../Apis/ChannelApi';
import TeamMemberPicker from "../Shared/TeamMemberPicker";
import { getGraphTokenUsingSsoToken } from "../../Apis/UtilApi";
import { BookingsBusiness, BookingsService, getBookingsBusinesses, getBookingsServices } from "../../Apis/GraphApi";
import { withTranslation, WithTranslation } from "react-i18next";
import { TaskModuleResult } from "../../Common/TaskModules";
import { AlertHandler } from "../../Models/AlertHandler";
import { DropdownItem } from "../../Models/ComponentItems";
import { TeamMember } from "../../Apis/AgentApi";

// route parameters
type RouteParams = {
    category: string;
}

// component properties
export interface AddEditServiceModalProps extends RouteComponentProps<RouteParams>, WithTranslation {
    alertHandler: AlertHandler;
}

// component state
export interface AddEditServiceModalState {
    appToken: string;
    graphToken: string;
    channels: Channel[];
    mappings: ChannelMapping[];
    mapping: Partial<ChannelMapping>;
    bookingsBusinesses: BookingsBusiness[];
    bookingsServices: BookingsService[];
    selectedTeamId?: string;
    errors: {
        category: boolean,
        categoryDuplicate: boolean,
        channel: boolean,
        business: boolean,
        service: boolean,
    };
    saving: boolean;
}

// AddEditServiceModal component
class AddEditServiceModal extends React.Component<AddEditServiceModalProps, AddEditServiceModalState> {
    constructor(props: AddEditServiceModalProps) {
        super(props);
        this.state = {
            appToken: "",
            graphToken: "",
            channels: [],
            mappings: [],
            mapping: {
                category: "",
                channelId: "",
                supervisors: [],
                bookingsBusiness: null,
                bookingsService: null,
            },
            errors: {
                category: false,
                categoryDuplicate: false,
                channel: false,
                business: false,
                service: false,
            },
            saving: false,
            bookingsBusinesses: [],
            bookingsServices: [],
        };

        // perform the SSO request
        const authTokenRequest = {
            successCallback: this.tokenCallback,
            failureCallback: (error: string) => { console.log("Failure: " + error); },
        };
        microsoftTeams.authentication.getAuthToken(authTokenRequest);
    }

    // token callback from getAuthToken
    tokenCallback = async (token: string) => {
        this.setState({ appToken: token });

        // get token to get user photos
        const graphToken = await getGraphTokenUsingSsoToken(token);
        this.setState({ graphToken: graphToken });

        // get channels and mappings in parallel
        const [channels, mappings, bookingsBusinesses] = await Promise.all([
            getChannels(token),
            getChannelMappings(token),
            getBookingsBusinesses(graphToken),
        ]);

        if (channels.length === 0) {
            const errors = this.state.errors;
            errors.channel = true;
            this.props.alertHandler(this.props.t('errorNoChannelFound'), "danger");
        }

        // set the selected mapping if this is an edit (category passed in)
        let mapping = this.state.mapping;
        let selectedTeamId = this.state.selectedTeamId;
        if (this.props.match.params.category) {
            // find the mapping
            const foundMapping = mappings.find(m => m.category === this.props.match.params.category);
            if (foundMapping) {
                mapping = foundMapping;
                const channel = channels.find(c => c.channelId === mapping.channelId);
                selectedTeamId = channel.teamAadObjectId;

                // load the services for the selected mapping
                if (mapping.bookingsBusiness) {
                    const bookingsServices = await getBookingsServices(this.state.graphToken, mapping.bookingsBusiness.id);
                    this.setState({ bookingsServices: bookingsServices });
                }
            }
        }

        // save objects to state
        this.setState({ mappings, channels, mapping, selectedTeamId, bookingsBusinesses });

        // show the UI
        microsoftTeams.appInitialization.notifyAppLoaded();
        microsoftTeams.appInitialization.notifySuccess();
    };

    // cancels the task module
    onCancel = () => {
        microsoftTeams.tasks.submitTask();
    };

    // saves the service
    onSave = async () => {
        let mapping = this.state.mapping;
        const errors = this.state.errors;
        let errorMsg = "";
        const isUpdate: boolean = (mapping.id && mapping.id !== "");

        // check if category populated
        if (!mapping.category || mapping.category.length === 0) {
            errors.category = true;
            errorMsg = this.props.t('errorCategoryRequired');
        }

        // check if the category is already in use
        console.log(mapping);
        const dups = this.state.mappings.filter((item) => {
            return item.category === mapping.category;
        });
        if ((!isUpdate && dups.length > 0) ||
            (isUpdate && dups.length > 1)) {
            errors.categoryDuplicate = true;
            errorMsg = errorMsg = this.props.t('errorCategoryInUse');
        }

        // check if channel selected
        if (!mapping.channelId || mapping.channelId.length === 0) {
            errors.channel = true;
            errorMsg = errorMsg + ((errorMsg.length > 0) ? "; " : "") + this.props.t('errorChannelRequired');
        }

        // check if bookings business is selected
        if (!mapping.bookingsBusiness || mapping.bookingsBusiness.id.length === 0) {
            errors.business = true;
            errorMsg = errorMsg + ((errorMsg.length > 0) ? "; " : "") + this.props.t('errorBookingBusinessRequired');
        }

        // check if bookings business is selected
        if (!mapping.bookingsService || mapping.bookingsService.id.length === 0) {
            errors.service = true;
            errorMsg = errorMsg + ((errorMsg.length > 0) ? "; " : "") + this.props.t('errorBookingServiceRequired');
        }

        // show error message
        if (this.state.errors.category ||
            this.state.errors.categoryDuplicate ||
            this.state.errors.channel ||
            this.state.errors.business ||
            this.state.errors.service) {
            this.props.alertHandler(errorMsg, "danger");
        }
        else {
            // disable the loading indicator
            this.setState({ saving: true });

            // get the mapping by editIndex and check if it is an add or update
            if (isUpdate) {
                // update the existing mapping
                await patchChannelMapping(this.state.appToken, mapping);
            }
            else {
                // create new mapping
                mapping = await postChannelMapping(this.state.appToken, mapping);
            }

            // close the task module, submitting the new/updated channel mapping as the result
            const taskModuleResult: TaskModuleResult = { type: 'channelMappingResult', channelMapping: mapping as ChannelMapping };
            microsoftTeams.tasks.submitTask(taskModuleResult);
        }
    };

    // gets the combined team and channel name
    getChannelDisplayName = (channel: Channel): string => {
        return `${channel.teamName} - ${channel.channelName}`;
    }

    // handles selected channel change
    selectedChannelChanged = (_evt: unknown, ctrl: DropdownProps) => {
        // update channelId of mapping
        const selectedItem = ctrl.value as DropdownItem<Channel>;
        this.setState(prevState => ({
            mapping: { ...prevState.mapping, channelId: selectedItem.data.channelId },
            selectedTeamId: selectedItem.data.teamAadObjectId,
            errors: { ...prevState.errors, channel: false },
        }));
    };

    // handles category change
    categoryChanged: ComponentEventHandler<InputProps & { value: string; }> = (_evt, ctrl) => {
        // update category of mapping
        this.setState(prevState => ({
            mapping: { ...prevState.mapping, category: ctrl.value },
            errors: { ...prevState.errors, category: ctrl.value.length === 0, categoryDuplicate: false },
        }));
    };

    // handles the bookings business change
    bookingsBusinessChanged = async (_evt: unknown, ctrl: DropdownProps) => {
        // update business of mapping
        const selectedItem = ctrl.value as DropdownItem<BookingsBusiness>;
        this.setState(prevState => ({
            mapping: { ...prevState.mapping, bookingsBusiness: selectedItem.data, bookingsService: null },
            bookingsServices: [],
            errors: { ...prevState.errors, business: false },
        }));

        const bookingsServices = await getBookingsServices(this.state.graphToken, selectedItem.data.id);
        this.setState({ bookingsServices });
    };

    // handles the bookings service change
    bookingsServiceChanged = (_evt: unknown, ctrl: DropdownProps) => {
        // update service of mapping
        const selectedItem = ctrl.value as DropdownItem<BookingsService>;
        this.setState(prevState => ({
            mapping: { ...prevState.mapping, bookingsService: selectedItem.data },
            errors: { ...prevState.errors, service: false },
        }));
    };

    // handles members changed event
    membersChanged = (members: TeamMember[]) => {
        this.setState(prevState => ({
            mapping: { ...prevState.mapping, supervisors: members },
        }));
    };

    private getChannelDropdownItem(channel: Channel): DropdownItem<Channel> {
        return channel ? { header: this.getChannelDisplayName(channel), key: channel.id, data: channel } : null;
    }

    private getBookingsBusinessDropdownItem(business: BookingsBusiness): DropdownItem<BookingsBusiness> {
        return business ? { header: business.displayName, key: business.id, data: business } : null;
    }

    private getBookingsServiceDropdownItem(service: BookingsService): DropdownItem<BookingsService> {
        return service ? { header: service.displayName, key: service.id, data: service } : null;
    }

    // renders the component
    render() {
        return (
            <div className="page" style={{ padding: "20px", display: "flex", flexDirection: "column", height: "100vh" }}>
                <Flex gap="gap.small" padding="padding.medium">
                    <Input label={this.props.t('serviceNameLabel')} error={this.state.errors.category} fluid value={this.state.mapping.category} onChange={this.categoryChanged.bind(this)} />
                </Flex>
                <Flex gap="gap.small" padding="padding.medium">
                    <div style={{ width: "100%" }}>
                        <Text content={this.props.t('channelLabel')} style={{ width: "100%" }} />
                        <Dropdown
                            items={this.state.channels.map(c => this.getChannelDropdownItem(c))}
                            value={this.getChannelDropdownItem(this.state.channels.find(c => c.channelId === this.state.mapping.channelId))}
                            placeholder={this.props.t('channelPlaceholder')}
                            checkable
                            fluid
                            error={this.state.errors.channel}
                            noResultsMessage={this.props.t('dropdownNoResults')}
                            onChange={this.selectedChannelChanged.bind(this)} />
                    </div>
                </Flex>
                <Flex gap="gap.small" padding="padding.medium">
                    <div style={{ width: "100%" }}>
                        <Text content={this.props.t('bookingsBusinessLabel')} style={{ width: "100%" }} />
                        <Dropdown
                            items={this.state.bookingsBusinesses.map(b => this.getBookingsBusinessDropdownItem(b))}
                            value={this.getBookingsBusinessDropdownItem(this.state.mapping.bookingsBusiness)}
                            placeholder={this.props.t('bookingsBusinessPlaceholder')}
                            checkable
                            fluid
                            error={this.state.errors.business}
                            noResultsMessage={this.props.t('dropdownNoResults')}
                            onChange={this.bookingsBusinessChanged.bind(this)} />
                    </div>
                </Flex>
                <Flex gap="gap.small" padding="padding.medium">
                    <div style={{ width: "100%" }}>
                        <Text content={this.props.t('bookingsServiceLabel')} style={{ width: "100%" }} />
                        <Dropdown
                            items={this.state.bookingsServices.map(s => this.getBookingsServiceDropdownItem(s))}
                            value={this.getBookingsServiceDropdownItem(this.state.mapping.bookingsService)}
                            placeholder={this.props.t('bookingsServicePlaceholder')}
                            disabled={this.state.bookingsServices.length < 1}
                            checkable
                            fluid
                            error={this.state.errors.service}
                            noResultsMessage={this.props.t('dropdownNoResults')}
                            onChange={this.bookingsServiceChanged.bind(this)} />
                    </div>
                </Flex>
                <Flex gap="gap.small" padding="padding.medium">
                    <div style={{ width: "100%" }}>
                        <Text content={this.props.t('supervisorsLabel')} style={{ width: "100%" }} />
                        <TeamMemberPicker
                            teamAadObjectId={this.state.selectedTeamId}
                            appToken={this.state.appToken}
                            graphToken={this.state.graphToken}
                            value={this.state.mapping.supervisors}
                            onChange={this.membersChanged.bind(this)}
                            placeholder={this.props.t('supervisorsPlaceholder')}
                            onError={this.props.alertHandler.bind(this)} />
                    </div>
                </Flex>
                <Flex style={{ flex: 1 }}></Flex>
                <Divider />
                <Flex hAlign="end" gap="gap.small">
                    <Button content={this.props.t('cancelButton')} onClick={() => this.onCancel()} />
                    <Button primary content={this.props.t('saveButton')} onClick={() => this.onSave()} loading={this.state.saving} disabled={this.state.saving} />
                </Flex>
            </div>
        );
    }
}

export default withTranslation(['addEditServiceModal', 'common'])(AddEditServiceModal);