import * as React from "react";
import {
    RouteComponentProps,
} from "react-router-dom";
import * as microsoftTeams from "@microsoft/teams-js";
import { Avatar, Text, Table, Flex, SplitButton, Button, Accordion, Input, ParticipantAddIcon, SearchIcon, ExclamationCircleIcon, Layout, MenuButton, VolumeDownIcon, ParticipantRemoveIcon, RedoIcon, MenuIcon, ComponentEventHandler, InputProps, ShorthandCollection, TableRowProps, MenuItemProps } from "@fluentui/react-northstar";
import TimeBlock from "../Shared/TimeBlock";

import { ConsultDetails, getFilteredRequests, RequestStatus } from '../../Apis/ConsultApi';
import { getGraphTokenUsingSsoToken } from "../../Apis/UtilApi";
import { assignSelfTaskModule, assignToOtherTaskModule, detailsTaskModule, TaskModuleResult } from "../../Common/TaskModules";
import { ConsultScheduleTabSettings } from "../../Models/ConsultScheduleTabSettings";
import RequestCard from "../Shared/RequestCard";
import { PhotoUtil } from "../../Utils/PhotoUtil";
import { Trans, withTranslation, WithTranslation } from "react-i18next";
import { AlertHandler } from "../../Models/AlertHandler";

// route parameters
type RouteParams = {
    config: string;
}

// component properties
export interface ConsultScheduleTabProps extends RouteComponentProps<RouteParams>, WithTranslation {
    alertHandler: AlertHandler;
}

// component state
export interface ConsultScheduleTabState {
    token: string;
    graphToken: string;
    context: microsoftTeams.Context;
    requests: ConsultDetails[];
    settings: ConsultScheduleTabSettings;
    search: string;
    isMobile: boolean;
    filter?: string;
}

// ConsultScheduleTab component
class ConsultScheduleTab extends React.Component<ConsultScheduleTabProps, ConsultScheduleTabState> {
    photoUtil: PhotoUtil = new PhotoUtil();
    constructor(props: ConsultScheduleTabProps) {
        super(props);
        this.state = {
            token: "",
            graphToken: "",
            context: null,
            requests: [],
            settings: JSON.parse(atob(this.props.match.params.config)),
            search: null,
            isMobile: false,
        };

        // authenticate the user
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
        // Store the token in state
        this.setState({ token: token });

        // get context so we can look up requests for a specific channel
        microsoftTeams.getContext(async (ctx: microsoftTeams.Context) => {
            const isMobile: boolean = (ctx.hostClientType === "ios" || ctx.hostClientType === "android");
            this.setState({ context: ctx, isMobile: isMobile });

            // fetch the requests
            const [requests, graphToken] = await Promise.all([getFilteredRequests(token, this.state.settings), getGraphTokenUsingSsoToken(token)]);
            this.setState({ requests: requests, graphToken: graphToken });
            microsoftTeams.appInitialization.notifySuccess();
        });
    }

    // launches the assign to me task module
    assignMe = (requestId: string) => {
        const assignSelfTaskInfo = assignSelfTaskModule(requestId, this.props.t);
        microsoftTeams.tasks.startTask(assignSelfTaskInfo, this.assignTaskModuleSubmitHandler);
    };

    // launches the assign to others task module
    assignOther = (requestId: string) => {
        const assignToOtherTaskInfo = assignToOtherTaskModule(requestId, this.props.t);
        microsoftTeams.tasks.startTask(assignToOtherTaskInfo, this.assignTaskModuleSubmitHandler);
    };

    assignTaskModuleSubmitHandler = (err: string, result: unknown) => {
        const tmResult = result as TaskModuleResult;
        if (!err && tmResult?.type === 'consultDetailsResult') {
            // reload the tab to refresh consult data
            this.setState(prevState => ({
                requests: prevState.requests.map(
                    r => r.id === tmResult.consultDetails.id ? tmResult.consultDetails : r
                ),
            }));
        }
    }

    // launches the online meeting
    joinCall = (request: ConsultDetails) => {
        if (request && request.joinUri && request.joinUri.length > 0) {
            microsoftTeams.executeDeepLink(request.joinUri);
        }
        else {
            this.props.alertHandler(this.props.t('errorCannotJoin'), "warning");
        }
    };

    // launches the view details task module
    openDetails = (request: ConsultDetails) => {
        const detailsTaskInfo = detailsTaskModule(request.id, this.props.t);
        microsoftTeams.tasks.startTask(detailsTaskInfo, (err, result: unknown) => {
            const tmResult = result as TaskModuleResult;
            if (!err && tmResult?.type === 'consultDetailsResult') {
                // reload the tab to refresh consult data
                this.setState(prevState => ({
                    requests: prevState.requests.map(
                        r => r.id === tmResult.consultDetails.id ? tmResult.consultDetails : r
                    ),
                }));
            }
        });
    };

    // applies a search filter on the displayed data
    onSearchFilter: ComponentEventHandler<InputProps & { value: string; }> = (_evt, data) => {
        this.setState({ search: data.value });
    };

    // event for filter changed
    onFilter: ComponentEventHandler<MenuItemProps> = (evt, ctrl) => {
        const filter: string = (ctrl.itemPosition === 1) ? "Unassigned" : ((ctrl.itemPosition === 2) ? "ReassignRequested" : ((ctrl.itemPosition === 3) ? "Assigned" : null));
        this.setState({ filter: filter });
    };

    // event raised when an action is executed on the item
    onAction = (item: ConsultDetails, action: string) => {
        switch (action) {
            case "detail":
                this.openDetails(item);
                break;
            case "assignMe":
                this.assignMe(item.id);
                break;
            case "assignOther":
                this.assignOther(item.id);
                break;
            case "requestReassign":
                break;
            case "markComplete":
                this.openDetails(item);
                break;
            case "join":
                this.joinCall(item);
                break;
        }
    };

    // renders the component
    render() {
        if (this.state.isMobile) {
            const items = this.state.requests.filter((request) => {
                // filter the view based on the search box and filter
                if (!this.state.filter || this.state.filter === "" || this.state.filter === request.status) {
                    if (!this.state.search || this.state.search.length < 2) {
                        return true;
                    } else if (request.friendlyId.toLowerCase().includes(this.state.search.toLowerCase()) ||
                        request.customerName.toLowerCase().includes(this.state.search.toLowerCase()) ||
                        request.category.toLowerCase().includes(this.state.search.toLowerCase()) ||
                        request.status.toLowerCase().includes(this.state.search.toLowerCase()) ||
                        request.query.toLowerCase().includes(this.state.search.toLowerCase())) {
                        return true;
                    }
                }

                return false;
            }).map((item) => {
                return (
                    <RequestCard request={item} graphToken={this.state.graphToken} meAadId={this.state.context.userObjectId} photoUtil={this.photoUtil} onAction={this.onAction} />
                );
            });

            return (
                <div style={{ padding: "20px" }}>
                    <Layout
                        main={<Input icon={<SearchIcon />} clearable inverted fluid placeholder={this.props.t("searchPlaceholder")} onChange={this.onSearchFilter.bind(this)} />}
                        end={<MenuButton
                            trigger={<Button iconOnly icon={<VolumeDownIcon style={{ transform: "rotate(-90deg)", marginTop: "-8px" }} size="large" />} title={this.props.t("openFilterButton")} />}
                            onMenuItemClick={this.onFilter.bind(this)}
                            menu={[
                                (<Flex>
                                    <ParticipantRemoveIcon style={{ color: "#B51C3B" }} />
                                    <Text style={{ paddingLeft: "10px" }} content={this.props.t("stateUnassigned")}></Text>
                                </Flex>),
                                (<Flex>
                                    <RedoIcon style={{ color: "#9F5811" }} />
                                    <Text style={{ paddingLeft: "10px" }} content={this.props.t("stateReassign")}></Text>
                                </Flex>),
                                (<Flex>
                                    <ParticipantAddIcon style={{ color: "#1F6C3D" }} />
                                    <Text style={{ paddingLeft: "10px" }} content={this.props.t("stateAssigned")}></Text>
                                </Flex>),
                                (<Flex>
                                    <MenuIcon />
                                    <Text style={{ paddingLeft: "10px" }} content={this.props.t("stateAll")}></Text>
                                </Flex>),
                            ]}
                        />}
                    />
                    {items}
                </div>
            );
        } else {
            // build table headers
            const header = {
                key: "header",
                items: [
                    { content: <Text weight="regular" content={this.props.t('idColHeader')} />, key: "id", style: { maxWidth: "90px", backgroundColor: "transparent" } },
                    { content: <Text weight="regular" content={this.props.t('nameColHeader')} />, key: "customerName" },
                    { content: <Text weight="regular" content={this.props.t('categoryColHeader')} />, key: "category", style: { maxWidth: "100px" } },
                    { content: <Text weight="regular" content={this.props.t('timeColHeader')} />, key: "assignedTimeBlock", style: { maxWidth: "250px" } },
                    { content: <Text weight="regular" content={this.props.t('statusColHeader')} />, key: "status" },
                    { content: <Text weight="regular" content="" />, key: "actions", style: { maxWidth: "140px" } },
                ],
            };

            // Build view based on assigned status of requests
            const unassignedRows: ShorthandCollection<TableRowProps> = [];
            const assignedRows: ShorthandCollection<TableRowProps> = [];
            let reassignmentCount = 0;
            this.state.requests.filter((request) => {
                // filter the view based on the search box
                if (this.state.search == null) {
                    return true;
                } else if (request.friendlyId.toLowerCase().includes(this.state.search.toLowerCase()) ||
                    request.customerName.toLowerCase().includes(this.state.search.toLowerCase()) ||
                    request.category.toLowerCase().includes(this.state.search.toLowerCase()) ||
                    request.status.toLowerCase().includes(this.state.search.toLowerCase())) {
                    return true;
                }

                return false;
            }).forEach((request, index) => {
                // build the status JSX
                let statusJSX;
                switch (request.status) {
                    case RequestStatus.Unassigned:
                        statusJSX = <Text error content={this.props.t('stateUnassigned')}></Text>;
                        break;
                    case RequestStatus.ReassignRequested:
                        statusJSX = <Flex><Text color="orange" content={this.props.t('stateReassign')}></Text><ExclamationCircleIcon color="orange" style={{ paddingLeft: "6px" }} /></Flex>;
                        reassignmentCount++;
                        break;
                    case RequestStatus.Assigned:
                        const isSelf = request.assignedToId === this.state.context.userObjectId;
                        statusJSX = (
                            // <Text> content will get populated from i18n translation file
                            <Flex>
                                <Trans i18nKey={isSelf ? "stateAssignedToYou" : "stateAssignedTo"} t={this.props.t} values={isSelf ? {} : { name: request.assignedToName }}>
                                    <Text />
                                    <Text weight="bold" style={{ paddingLeft: "6px" }} />
                                </Trans>
                            </Flex>
                        );
                        break;
                    case RequestStatus.Completed:
                        statusJSX = <Text success content={this.props.t('stateCompleted')}></Text>;
                        break;
                    default:
                        statusJSX = <Text content={request.status}></Text>;
                }

                // create the row
                const row = {
                    key: index,
                    items: [
                        {
                            content: <Button content={request.friendlyId} text primary className="textOnly" onClick={() => this.openDetails(request)} />,
                            style: { maxWidth: "90px" },
                        },
                        {
                            content:
                                <Flex>
                                    <Avatar name={request.customerName} />
                                    <span style={{ paddingLeft: "10px", paddingTop: "6px" }}>{request.customerName}</span>
                                </Flex>,
                        },
                        {
                            content: <Text content={request.category}></Text>, style: { maxWidth: "100px" },
                        },
                        {
                            content: request.status !== RequestStatus.Unassigned
                                ? <TimeBlock timeBlock={request.assignedTimeBlock}></TimeBlock>
                                : null,
                            style: { maxWidth: "250px" },
                        },
                        {
                            content: statusJSX,
                        },
                        {
                            content:
                                (request.status === RequestStatus.Unassigned) ? (
                                    <SplitButton
                                        menu={[
                                            {
                                                key: "me",
                                                content: this.props.t('assignMeButton'),
                                                icon: <ParticipantAddIcon />,
                                                onClick: () => { this.assignMe(request.id); },
                                            },
                                            {
                                                key: "other",
                                                content: this.props.t('assignOtherButton'),
                                                icon: <ParticipantAddIcon />,
                                                onClick: () => { this.assignOther(request.id); },
                                            },
                                        ]}
                                        button={{
                                            content: this.props.t('assignButton'),
                                            "aria-roledescription": "splitbutton",
                                            "aria-describedby": "instruction-message-primary-button",
                                        }}
                                        primary
                                        toggleButton={{
                                            "aria-label": "more options",
                                        }}
                                        onMainButtonClick={() => this.assignOther(request.id)}
                                    />
                                ) : ((request.status === RequestStatus.Assigned) ? (
                                    <Button content={this.props.t('joinButton')} iconPosition="before" primary onClick={() => this.joinCall(request)} />
                                ) : (
                                    <Button content={this.props.t('detailsButton')} iconPosition="before" onClick={() => this.openDetails(request)} />
                                )),
                            style: { maxWidth: "140px", display: "flex", justifyContent: "flex-end" },
                        },
                    ],
                };

                // add the row to the correct rows collection based on assigned status
                if (request.status === RequestStatus.Unassigned) {
                    unassignedRows.push(row);
                } else {
                    assignedRows.push(row);
                }
            });

            // build the panels for each status
            const reassignCnt = (reassignmentCount > 0) ? (<Text content={this.props.t('reassignCount', { count: reassignmentCount })} color="orange" style={{ paddingLeft: "6px" }} />) : (<Text content="" />);
            const panels = [
                {
                    title: (
                        // <Text> content will get populated from i18n translation file
                        <Flex>
                            <Trans i18nKey="unassignedSection" t={this.props.t} count={unassignedRows.length}>
                                <Text weight="bold" />
                                <Text style={{ paddingLeft: "6px" }} />
                            </Trans>
                        </Flex>
                    ),
                    content: (<Table compact header={header} rows={unassignedRows} aria-label="Compact view static table" />),
                },
                {
                    title: (
                        // <Text> content will get populated from i18n translation file
                        <Flex>
                            <Trans i18nKey="assignedSection" t={this.props.t} count={assignedRows.length}>
                                <Text weight="bold" />
                                <Text style={{ paddingLeft: "6px" }} />
                            </Trans>
                            {reassignCnt}
                        </Flex>
                    ),
                    content: (<Table compact header={header} rows={assignedRows} aria-label="Compact view static table" />),
                },
            ];

            return (
                <div className="page">
                    <Flex hAlign="end">
                        <Input icon={<SearchIcon />} clearable inverted placeholder={this.props.t('searchPlaceholder')} onChange={this.onSearchFilter.bind(this)} />
                    </Flex>
                    <Accordion defaultActiveIndex={[0]} panels={panels} />
                </div>
            );
        }
    }
}

export default withTranslation(['consultScheduleTab', 'common'])(ConsultScheduleTab);