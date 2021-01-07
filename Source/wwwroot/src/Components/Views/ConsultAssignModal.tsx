import * as React from "react";
import { RouteComponentProps } from "react-router-dom";
import * as microsoftTeams from "@microsoft/teams-js";
import { Header, Flex, Text, Divider, Layout, Avatar, Loader } from "@fluentui/react-northstar";
import { withTranslation, WithTranslation } from 'react-i18next';
import moment from "moment";

// local imports
import { assignConsult, getAvailability, getConsultDetails, ConsultDetails, MeetingDetail, getMeetingDetails, RequestStatus, isSupervisor } from "../../Apis/ConsultApi";
import { getTeamMembers } from "../../Apis/AgentApi";
import { AgentAvailability } from "../../Models/AgentAvailability";
import { TimeBlock as TB } from "../../Models/TimeBlock";
import TimeBlock from "../Shared/TimeBlock";
import { getGraphTokenUsingSsoToken } from "../../Apis/UtilApi";
import { PhotoUtil } from "../../Utils/PhotoUtil";
import AssignmentBlocked from "../Shared/AssignmentBlocked";
import AgentSelection from "../Shared/AgentSelection";
import TimeBlockSelection from "../Shared/TimeBlockSelection";
import AssignmentComments from "../Shared/AssignmentComments";
import AssignmentActions from "../Shared/AssignmentActions";
import { AssignmentAction, BlockedReason, AssignmentStep } from "../../Models/AssignmentEnums";
import { TaskModuleResult } from "../../Common/TaskModules";

// route parameters
type RouteParams = {
    requestId: string;
    mode: ConsultAssignModalMode;
}

// location state
export type ConsultAssignModalLocationState = {
    request: ConsultDetails;
}

// the modes that this modal can run in
type ConsultAssignModalMode = "self" | "other";

// component properties
export interface ConsultAssignModalProps extends RouteComponentProps<RouteParams, unknown, ConsultAssignModalLocationState>, WithTranslation {
    alertHandler: (alert: string, alertType: string) => void;
}

// component state
export interface ConsultAssignModalState {
    token: string;
    graphToken: string;
    consultDetails: ConsultDetails;
    availability: AgentAvailability[];
    context?: microsoftTeams.Context;
    blockReason: BlockedReason;
    selectedAgent?: AgentAvailability;
    selectedTime?: TB;
    comment: string;
    breadcrumb: AssignmentStep[];
    expandedTimePrefs: TB[];
    assignedPhoto: string;
    meetingDetails: MeetingDetail[];
    saving: boolean;
    loaded: boolean;
}

// ConsultAssignModal component
class ConsultAssignModal extends React.Component<ConsultAssignModalProps, ConsultAssignModalState> {
    constructor(props: ConsultAssignModalProps) {
        super(props);
        this.state = {
            token: null,
            graphToken: null,
            consultDetails: null,
            availability: [],
            blockReason: BlockedReason.NotBlocked,
            comment: "",
            breadcrumb: [AssignmentStep.SelectTime],
            expandedTimePrefs: [],
            assignedPhoto: "",
            meetingDetails: null,
            saving: false,
            loaded: false,
        };
    }

    // initialize photo util for profile photos
    photoUtil: PhotoUtil = new PhotoUtil();

    // component mounted start SSO auth flow
    componentDidMount() {
        // perform the SSO request
        const authTokenRequest: microsoftTeams.authentication.AuthTokenRequest = {
            successCallback: this.tokenCallback,
            failureCallback: this.tokenFailedCallback,
        };
        microsoftTeams.authentication.getAuthToken(authTokenRequest);

        // initialize context
        microsoftTeams.getContext((ctx: microsoftTeams.Context) => {
            this.setState({ context: ctx });
        });
    }

    // token callback for SSO
    tokenCallback = async (token: string) => {
        this.setState({ token });

        // Get the graph token
        const graphToken = await getGraphTokenUsingSsoToken(token);

        // Check if this is assign to another agent and if user can do that
        let breadcrumb = this.state.breadcrumb;
        const isOther = this.props.match.params.mode === "other";
        let canSelectOther = false;
        if (isOther) {
            breadcrumb = [AssignmentStep.SelectAgent];
            canSelectOther = await isSupervisor(token, this.props.match.params.requestId);
        }

        // Get consult details
        const consultDetails = this.props.location.state?.request ?? await getConsultDetails(token, this.props.match.params.requestId);
        const expandedTimePrefs = this.expandPreferredTimes(consultDetails);

        // Get availability for self or team based on mode AND pass in time block(s) based on status
        const blocks: TB[] = (consultDetails.status === RequestStatus.Unassigned) ? consultDetails.preferredTimes : [consultDetails.assignedTimeBlock];
        let availability = await getAvailability(token, blocks, (isOther && canSelectOther) ? this.state.context.groupId : null);

        // Determine if there is availability
        const busy = !availability.some(a => a.timeBlocks.length > 0);

        // Check for a blocked reason
        let blockedReason: BlockedReason = BlockedReason.NotBlocked;
        if (consultDetails.status === RequestStatus.Assigned) {
            // This consult is already assigned
            blockedReason = BlockedReason.AlreadyAssigned;
        }
        if (this.props.match.params.mode === "self" && busy) {
            // You are not available...override?
            blockedReason = BlockedReason.NoAvailabilitySelf;
        }
        else if (this.props.match.params.mode === "other" && !canSelectOther) {
            // You are not an authorized to assign others...assign self?
            blockedReason = (busy) ? BlockedReason.NotAuthorizedAndNoAvailability : BlockedReason.NotAuthorized;
        }
        else if (this.props.match.params.mode === "other" && canSelectOther && busy) {
            // No agents are available...override?
            blockedReason = BlockedReason.NoAvailabilityTeam;
            const teamMembers = await getTeamMembers(token, this.state.context.groupId);
            availability = teamMembers as AgentAvailability[];
            availability.forEach((agent: AgentAvailability) => {
                agent.timeBlocks = [];
            });
        }

        // get details if this is a reassign and blocked
        let meetingDetails: MeetingDetail[] = null;
        if (consultDetails.status === RequestStatus.ReassignRequested && busy) {
            meetingDetails = await getMeetingDetails(token, consultDetails.assignedTimeBlock);
        }

        // set the start view
        breadcrumb = (blockedReason !== BlockedReason.NotBlocked) ? [AssignmentStep.Blocked] : ((consultDetails.status === RequestStatus.ReassignRequested && !isOther) ? [AssignmentStep.Comments] : breadcrumb);

        // get assigned user photo if self
        if (this.props.match.params.mode === "self") {
            this.photoUtil.getGraphPhoto(graphToken, availability[0].id).then((img: string) => {
                this.setState({ assignedPhoto: img });
            });
        }

        // save everything in state
        this.setState({
            graphToken: graphToken,
            consultDetails: consultDetails,
            availability: availability,
            blockReason: blockedReason,
            selectedAgent: (this.props.match.params.mode === "self" || !canSelectOther) ? availability[0] : null,
            selectedTime: (consultDetails.status === RequestStatus.ReassignRequested) ? consultDetails.assignedTimeBlock : null,
            breadcrumb: breadcrumb,
            expandedTimePrefs: expandedTimePrefs,
            assignedPhoto: this.photoUtil.emptyPic,
            meetingDetails: meetingDetails,
            loaded: true,
        });

        // Stop spinner for component to display
        microsoftTeams.appInitialization.notifySuccess();
    };

    // token failure callback for SSO
    tokenFailedCallback = (error: string) => {
        console.error(`SSO failed: ${error}`);
        microsoftTeams.appInitialization.notifyFailure({
            reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
            message: this.props.t("errorAuthFailed"),
        });
    };

    // handles action buttons at the bottom of the task module
    actionClicked = async (action: AssignmentAction) => {
        const breadcrumb = this.state.breadcrumb;
        if (action === AssignmentAction.NoCancel || action === AssignmentAction.Cancel) {
            // Cancel/close the task module
            microsoftTeams.tasks.submitTask();
        }
        else if (action === AssignmentAction.Override) {
            // determine where to go next
            console.log(this.state.consultDetails.status);
            if (this.state.consultDetails.status === RequestStatus.Unassigned) {
                if (this.props.match.params.mode === "self") {
                    breadcrumb.push(AssignmentStep.SelectTime);
                }
                else if (this.state.blockReason === BlockedReason.NotAuthorized || this.state.blockReason === BlockedReason.NotAuthorizedAndNoAvailability) {
                    breadcrumb.push(AssignmentStep.SelectTime);
                }
                else {
                    breadcrumb.push(AssignmentStep.SelectAgent);
                }
            }
            else {
                if (this.props.match.params.mode === "self" || this.state.blockReason === BlockedReason.NotAuthorized || this.state.blockReason === BlockedReason.NotAuthorizedAndNoAvailability) {
                    breadcrumb.push(AssignmentStep.Comments);
                }
                else {
                    breadcrumb.push(AssignmentStep.SelectAgent);
                }
            }
            this.setState({ breadcrumb });
        }
        else if (action === AssignmentAction.Back) {
            // pop the last item off the breadcrumb to go back
            breadcrumb.pop();
            this.setState({ breadcrumb: breadcrumb });
        }
        else if (action === AssignmentAction.Assign) {
            this.setState({ saving: true });
            let agent: AgentAvailability = null;
            if (this.props.match.params.mode === "other") {
                agent = this.state.selectedAgent;
            }
            // Assign to self
            let consultDetails: ConsultDetails = null;
            try {
                consultDetails = await assignConsult(this.state.token,
                    this.state.consultDetails.id,
                    this.state.selectedTime,
                    this.state.comment,
                    agent);
            } catch (err) {
                this.setState({ saving: false });
                this.props.alertHandler(this.props.t('errorCannotAssign'), 'danger');
                return;
            }

            // close task module
            const taskModuleResult: TaskModuleResult = { type: 'consultDetailsResult', consultDetails };
            microsoftTeams.tasks.submitTask(taskModuleResult);
        }
    };

    // expands all preferred times into an expanded list of 30min blocks
    expandPreferredTimes = (consultDetails: ConsultDetails) => {
        const blocks: TB[] = [];
        consultDetails.preferredTimes.forEach((block: TB) => {
            let currMoment = moment(block.startDateTime);
            const endMoment = moment(block.endDateTime);
            while (endMoment.diff(currMoment, 'minute', true) >= 30) {
                const startDateTime = currMoment.format();
                currMoment = currMoment.add(30, "minutes");
                const endDateTime = currMoment.format();
                blocks.push({ startDateTime, endDateTime });
            }
        });
        return blocks;
    };

    // get assignment step text
    getHeader = (step: AssignmentStep) => {
        switch (step) {
            case AssignmentStep.Blocked:
                if (this.state.blockReason === BlockedReason.NotAuthorized || this.state.blockReason === BlockedReason.NotAuthorizedAndNoAvailability) {
                    return this.props.t("unauthorized");
                } else if (this.state.blockReason === BlockedReason.NoAvailabilityTeam) {
                    return this.props.t("selectAgentLabel");
                } else if (this.state.blockReason === BlockedReason.NoAvailabilitySelf) {
                    return this.props.t("selectTimeHeader");
                } else if (this.state.blockReason === BlockedReason.AlreadyAssigned) {
                    return this.props.t("errorAlreadyAssigned");
                }
                return null;
            case AssignmentStep.SelectAgent:
                return this.props.t("selectAgentLabel");
            case AssignmentStep.SelectTime:
                return this.props.t("selectTimeHeader");
            case AssignmentStep.Comments:
                return this.props.t("commentsLabel");
        }
    };

    // Comment changed event
    commentChanged = (text: string) => {
        this.setState({ comment: text });
    };

    // renders preferred time slot list
    private renderPreferredTimeBlockGroups(preferredTimesGroupedByDate: Record<string, TB[]>) {
        return Object.entries(preferredTimesGroupedByDate).map(([_, timeBlocks]) => {
            const dateStr = moment(timeBlocks[0].startDateTime).format('ll');
            const timeBlockStrs = timeBlocks.map(timeBlock => {
                const startTimeStr = moment(timeBlock.startDateTime).format('LT');
                const endTimeStr = moment(timeBlock.endDateTime).format('LT');
                return `${startTimeStr} - ${endTimeStr}`;
            });
            const timeStr = timeBlockStrs.join(", ");
            return (
                <Flex gap="gap.smaller">
                    <Text content={dateStr} weight="bold" />
                    <Text content={timeStr} weight="light" />
                </Flex>
            );
        });
    }

    // Renders the component
    render = () => {
        if (!this.state.loaded) {
            return (
                <div className="taskModule">
                    <div className="tmBody">
                        <Loader style={{ height: "100%" }} />
                    </div>
                </div>
            );
        }

        let ctrl = (<></>);
        let actions: AssignmentAction[] = [];
        if (this.state.breadcrumb[this.state.breadcrumb.length - 1] === AssignmentStep.Blocked) {
            // display blocked message
            ctrl = <AssignmentBlocked reason={this.state.blockReason} meetingDetail={this.state.meetingDetails} />;
            actions = [AssignmentAction.NoCancel, AssignmentAction.Override];
        }
        else if (this.state.availability && this.state.breadcrumb[this.state.breadcrumb.length - 1] === AssignmentStep.SelectAgent) {
            // allow user to select an agent
            ctrl = <AgentSelection
                agents={this.state.availability}
                graphToken={this.state.graphToken}
                photoUtil={this.photoUtil}
                onAgentSelectionChanged={(agent: AgentAvailability) => {
                    const breadcrumb = this.state.breadcrumb;
                    if (this.state.consultDetails.status === RequestStatus.ReassignRequested) {
                        breadcrumb.push(AssignmentStep.Comments);
                    } else {
                        breadcrumb.push(AssignmentStep.SelectTime);
                    }
                    this.photoUtil.getGraphPhoto(this.state.graphToken, agent.id).then((img: string) => {
                        this.setState({ assignedPhoto: img });
                    });
                    this.setState({
                        selectedAgent: agent,
                        breadcrumb: breadcrumb,
                    });
                }} />;

            if (this.state.breadcrumb.length > 1) {
                actions.push(AssignmentAction.Back);
            }
            actions.push(AssignmentAction.Cancel);
        }
        else if (this.state.selectedAgent && this.state.breadcrumb[this.state.breadcrumb.length - 1] === AssignmentStep.SelectTime) {
            // allow user to select a time block
            const blocks = (this.state.selectedAgent.timeBlocks.length > 0) ? this.state.selectedAgent.timeBlocks : this.state.expandedTimePrefs;
            ctrl = <TimeBlockSelection
                timeBlocks={blocks}
                onTimeBlockSelectionChanged={(tb: TB) => {
                    const breadcrumb = this.state.breadcrumb;
                    breadcrumb.push(AssignmentStep.Comments);
                    this.setState({
                        selectedTime: tb,
                        breadcrumb: breadcrumb,
                    });
                }} />;

            if (this.state.breadcrumb.length > 1) {
                actions.push(AssignmentAction.Back);
            }
            actions.push(AssignmentAction.Cancel);
        }
        else {
            // allow user to enter comments and confirm
            ctrl = <AssignmentComments comment={this.state.comment} commentChanged={this.commentChanged} />;

            if (this.state.breadcrumb.length > 1) {
                actions.push(AssignmentAction.Back);
            }
            actions.push(AssignmentAction.Cancel);
            actions.push(AssignmentAction.Assign);
        }

        // get time details
        let tbSection = <></>;
        if (this.state.consultDetails) {
            if (this.state.breadcrumb[this.state.breadcrumb.length - 1] === AssignmentStep.Comments) {
                tbSection = (
                    <>
                        <Text content={this.props.t("assignmentSummaryLabel")} size="small" className="tmSectionTitle" />
                        <div className="boxed selected">
                            <Layout start={<Avatar image={this.state.assignedPhoto} name={this.state.selectedAgent.displayName} />}
                                main={<Text content={this.state.selectedAgent.displayName} weight="bold" style={{ paddingLeft: "15px", paddingTop: "6px" }}></Text>}
                                end={<TimeBlock timeBlock={this.state.selectedTime} />}
                                style={{ width: "100%" }} />
                        </div>
                    </>
                );
            }
            else if (this.state.consultDetails.status === RequestStatus.ReassignRequested) {
                // show selected time
                tbSection = (
                    <div style={{ marginBottom: "1em" }}>
                        <Text content={this.props.t("scheduledTimeLabel")} size="small" className="tmSectionTitle" />
                        <TimeBlock timeBlock={this.state.consultDetails.assignedTimeBlock} />
                    </div>
                );
            }
            else {
                const preferredTimesSorted = [...this.state.consultDetails.preferredTimes]
                    .sort((a, b) => new Date(a.startDateTime).getTime() - new Date(b.startDateTime).getTime());
                const preferredTimesGroupedByDate = preferredTimesSorted.reduce<Record<string, TB[]>>((result, currentValue) => {
                    const startDateStr = new Date(currentValue.startDateTime).toDateString();
                    (result[startDateStr] = result[startDateStr] || []).push(currentValue);
                    return result;
                }, {});

                // show preferred times
                tbSection = (
                    <div style={{ marginBottom: "1em" }}>
                        <Text content={this.props.t("preferredTimesLabel")} size="small" className="tmSectionTitle" />
                        {this.state.consultDetails
                            && this.renderPreferredTimeBlockGroups(preferredTimesGroupedByDate)}
                        <br />
                    </div>
                );
            }
        }

        return (
            <div className="taskModule">
                <div className="tmBody">
                    <Header as="h3" content={this.getHeader(this.state.breadcrumb[this.state.breadcrumb.length - 1])} className="tmHeading" />
                    {tbSection}
                    {ctrl}
                </div>
                <div className="footer">
                    <Divider className="tmDivider" />
                    <AssignmentActions actions={actions} actionClicked={this.actionClicked} saving={this.state.saving} />
                </div>
            </div>
        );
    };
}

export default withTranslation(["consultAssignModal", "common"])(ConsultAssignModal);