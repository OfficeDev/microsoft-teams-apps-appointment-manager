import * as React from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { Table, Button, Text, Flex, Avatar, MenuButton, AcceptIcon, MoreIcon, ParticipantRemoveIcon, ComponentEventHandler, ButtonProps, Dialog, Layout, SearchIcon, Input, VolumeDownIcon, RedoIcon, ParticipantAddIcon, MenuIcon, MenuItemProps, ShorthandCollection, TableRowProps, InputProps } from '@fluentui/react-northstar';
import TimeBlock from "../Shared/TimeBlock";

import { getMyRequests, ConsultDetails, RequestStatus, completeConsult } from '../../Apis/ConsultApi';
import { detailsTaskModule, reassignTaskModule, TaskModuleResult } from "../../Common/TaskModules";
import { getGraphTokenUsingSsoToken } from "../../Apis/UtilApi";
import RequestCard from "../Shared/RequestCard";
import { PhotoUtil } from "../../Utils/PhotoUtil";

import { Trans, withTranslation, WithTranslation } from "react-i18next";
import { AlertHandler } from "../../Models/AlertHandler";

// component properties
export interface MyConsultsTabProps extends WithTranslation {
    alertHandler: AlertHandler;
}

// component state
export interface MyConsultsTabState {
    token: string;
    context: microsoftTeams.Context;
    isCompleteInProgress: boolean;
    requests: ConsultDetails[];
    completeRequestId: string;
    needCompleteConfirmation: boolean;
    isMobile: boolean;
    graphToken: string;
    filter?: string;
    search: string;
}

// MyConsultsTab component
class MyConsultsTab extends React.Component<MyConsultsTabProps, MyConsultsTabState> {
    photoUtil: PhotoUtil = new PhotoUtil();
    constructor(props: MyConsultsTabProps) {
        super(props);
        this.state = {
            token: "",
            graphToken: "",
            isCompleteInProgress: false,
            context: null,
            requests: [],
            needCompleteConfirmation: false,
            completeRequestId: null,
            isMobile: false,
            search: null,
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
        // save the access token into state
        this.setState({ token: token });

        // get teams context
        microsoftTeams.getContext(async (ctx: microsoftTeams.Context) => {
            const isMobile: boolean = (ctx.hostClientType === "ios" || ctx.hostClientType === "android");
            this.setState({ context: ctx, isMobile: isMobile });

            // get the user's consult requests
            const [requests, graphToken] = await Promise.all([getMyRequests(token), getGraphTokenUsingSsoToken(token)]);
            this.setState({ requests: requests, graphToken: graphToken });
            microsoftTeams.appInitialization.notifySuccess();
        });
    }

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

    // fires when a menu item is selected for a request item
    menuItemClick = (request: ConsultDetails, _evt: unknown, ctrl: MenuItemProps) => {
        // determine if this is a reassignment request or mark complete
        if (ctrl.index === 0) {
            // this is reassignment
            const reassignTaskInfo = reassignTaskModule(request.id, this.props.t);
            microsoftTeams.tasks.startTask(reassignTaskInfo, (err: string, result: unknown) => {
                const tmResult = result as TaskModuleResult;
                if (!err && tmResult?.type === 'consultDetailsResult') {
                    this.setState(prevState => ({
                        requests: prevState.requests.map(
                            r => r.id === tmResult.consultDetails.id ? tmResult.consultDetails : r
                        ),
                    }));
                }
            });
        }
        else {
            // this is mark complete
            this.setState({ needCompleteConfirmation: true, completeRequestId: request.id });
        }
    };

    // launches the online meeting
    joinCall = (request: ConsultDetails) => {
        if (request && request.joinUri && request.joinUri.length > 0) {
            microsoftTeams.executeDeepLink(request.joinUri);
        }
        else {
            this.props.alertHandler(this.props.t('errorCannotJoin'), "warning");
        }
    };

    // handler when mark complete, cancel button is clicked.
    cancelClicked: ComponentEventHandler<ButtonProps> = () => {
        this.setState({ needCompleteConfirmation: false, completeRequestId: null });
    };

    // handler when mark complete, confirm button is clicked.
    completeClicked: ComponentEventHandler<ButtonProps> = async () => {
        this.setState({ isCompleteInProgress: true });
        try {
            const consultDetails = await completeConsult(this.state.token, this.state.completeRequestId);
            this.setState(prevState => ({
                requests: prevState.requests.map(
                    r => r.id === consultDetails.id ? consultDetails : r
                ),
            }));
        } catch {
            this.setState({ isCompleteInProgress: false });
            this.props.alertHandler(this.props.t('errorCannotComplete'), "danger");
        } finally {
            this.setState({ needCompleteConfirmation: false, isCompleteInProgress: false, completeRequestId: null });
        }
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
                break;
            case "assignOther":
                break;
            case "requestReassign":
                this.menuItemClick(item, null, { index: 0 });
                break;
            case "markComplete":
                this.openDetails(item);
                break;
            case "join":
                this.joinCall(item);
                break;

        }
    };

    // applies search filter
    onSearchFilter: ComponentEventHandler<InputProps & { value: string; }> = (_evt, data) => {
        this.setState({ search: data.value });
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
            const header = {
                key: "header",
                items: [
                    { content: <Text weight="regular" content={this.props.t('idColHeader')} />, key: "id", style: { maxWidth: "90px", backgroundColor: "transparent" } },
                    { content: <Text weight="regular" content={this.props.t('nameColHeader')} />, key: "customerName" },
                    { content: <Text weight="regular" content={this.props.t('categoryColHeader')} />, key: "category", style: { maxWidth: "100px" } },
                    { content: <Text weight="regular" content={this.props.t('timeColHeader')} />, key: "assignedTimeBlock", style: { maxWidth: "250px" } },
                    { content: <Text weight="regular" content={this.props.t('statusColHeader')} />, key: "status" },
                    { content: <Text weight="regular" content="" />, key: "actions", style: { maxWidth: "160px" } },
                ],
            };

            const rows: ShorthandCollection<TableRowProps> = [];
            this.state.requests.forEach((request, index) => {
                let statusJSX;
                switch (request.status) {
                    case RequestStatus.Assigned:
                        statusJSX = (
                            // <Text> content will get populated from i18n translation file
                            <Flex>
                                <Trans i18nKey="stateAssignedToYou" t={this.props.t}>
                                    <Text />
                                    <Text weight="bold" style={{ paddingLeft: "6px" }} />
                                </Trans>
                            </Flex>
                        );
                        break;
                    case RequestStatus.Completed:
                        statusJSX = <Text success content={this.props.t('stateCompleted')}></Text>;
                        break;
                    case RequestStatus.ReassignRequested:
                        statusJSX = <Text content={this.props.t('stateReassign')}></Text>;
                        break;
                    default:
                        statusJSX = <Text content={request.status}></Text>;
                }
                rows.push({
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
                            content: <TimeBlock timeBlock={request.assignedTimeBlock}></TimeBlock>, style: { maxWidth: "250px" },
                        },
                        {
                            content: statusJSX,
                        },
                        {
                            content:
                                ((request.status === RequestStatus.Assigned) ? (
                                    <Flex>
                                        <Button content={this.props.t('joinButton')} iconPosition="before" primary onClick={() => this.joinCall(request)} />
                                        <MenuButton
                                            className="ellipse"
                                            trigger={<Button icon={<MoreIcon />} title={this.props.t('moreActionsButton')} />}
                                            menu={[
                                                {
                                                    index: 0,
                                                    content: (
                                                        <Flex>
                                                            <ParticipantRemoveIcon />
                                                            <Text style={{ paddingLeft: "10px" }} content={this.props.t('reassignButton')}></Text>
                                                        </Flex>
                                                    ),
                                                },
                                                {
                                                    index: 1,
                                                    content: (
                                                        <Flex>
                                                            <AcceptIcon />
                                                            <Text style={{ paddingLeft: "10px" }} content={this.props.t('completeButton')}></Text>
                                                        </Flex>
                                                    ),
                                                },
                                            ]}
                                            onMenuItemClick={this.menuItemClick.bind(this, request)}
                                        ></MenuButton>
                                    </Flex>
                                ) : (
                                    <div></div>
                                )), style: { maxWidth: "160px" },
                        },
                    ],
                });
            });
            return (
                <div className="page">
                    <Table compact header={header} rows={rows} aria-label="Compact view static table" />
                    <Dialog
                        open={this.state.needCompleteConfirmation || this.state.isCompleteInProgress}
                        header={this.props.t('completeConfirm')}
                        confirmButton={{ content: this.props.t('confirmButton'), loading: this.state.isCompleteInProgress, disabled: this.state.isCompleteInProgress }}
                        cancelButton={{ content: this.props.t('cancelButton'), disabled: this.state.isCompleteInProgress }}
                        onCancel={this.cancelClicked}
                        onConfirm={this.completeClicked}
                    />
                </div>
            );
        }
    }
}

export default withTranslation(['myConsultsTab', 'common'])(MyConsultsTab);