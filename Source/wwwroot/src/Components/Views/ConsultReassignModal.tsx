import * as React from "react";
import { RouteComponentProps } from "react-router-dom";
import * as microsoftTeams from "@microsoft/teams-js";
import { Button, Divider, Flex, TextArea, Text, ComponentEventHandler, ButtonProps, TextAreaProps } from '@fluentui/react-northstar';
import { ConsultDetails, getConsultDetails, getPostedChannel, reassignConsult, RequestStatus } from "../../Apis/ConsultApi";
import { getGraphTokenUsingSsoToken } from "../../Apis/UtilApi";
import TeamMemberPicker from "../Shared/TeamMemberPicker";
import { withTranslation, WithTranslation } from "react-i18next";
import { TaskModuleResult } from "../../Common/TaskModules";
import { AlertHandler } from "../../Models/AlertHandler";
import { TeamMember } from "../../Apis/AgentApi";

// route parameters
type RouteParams = {
    // the ID of the specific consult request
    requestId: string;
}

// location state
export type ConsultReassignModalLocationState = {
    request: ConsultDetails;
}

export interface ConsultReassignModalProps extends RouteComponentProps<RouteParams, unknown, ConsultReassignModalLocationState>, WithTranslation {
    alertHandler: AlertHandler;
}

export interface ConsultReassignModalState {
    token: string;
    graphToken: string;
    channelId: string;
    selectedAgents: TeamMember[];
    comment: string;
    selectedTeamId: string;
    saving: boolean;
    canReassign: boolean;
}

// ConsultReassignModal component
class ConsultReassignModal extends React.Component<ConsultReassignModalProps, ConsultReassignModalState> {
    constructor(props: ConsultReassignModalProps) {
        super(props);
        this.state = {
            token: null,
            graphToken: null,
            channelId: null,
            selectedAgents: [],
            comment: "",
            selectedTeamId: "",
            saving: false,
            canReassign: false,
        };
    }

    componentDidMount() {
        // perform the SSO request
        const authTokenRequest: microsoftTeams.authentication.AuthTokenRequest = {
            successCallback: this.tokenCallback,
            failureCallback: this.tokenFailedCallback,
        };
        microsoftTeams.authentication.getAuthToken(authTokenRequest);
    }

    // token callback from for SSO
    tokenCallback = async (token: string) => {
        // save token to state
        this.setState({ token: token });

        // get the graph token, channel details, and consult details in parallel
        const [graphToken, channel, consultDetails] = await Promise.all([
            getGraphTokenUsingSsoToken(token),
            getPostedChannel(token, this.props.match.params.requestId),
            this.props.location.state?.request
                ? Promise.resolve(this.props.location.state.request)
                : getConsultDetails(token, this.props.match.params.requestId),
        ]);

        if (consultDetails.status === RequestStatus.Assigned) {
            this.setState({ canReassign: true });
        } else if (consultDetails.status === RequestStatus.ReassignRequested) {
            this.props.alertHandler(this.props.t('errorAlreadyRequested'), 'danger');
        } else {
            this.props.alertHandler(this.props.t('errorCannotReassign'), 'danger');
        }

        // save state and stop spinner
        this.setState({ graphToken: graphToken });
        this.setState({ selectedTeamId: channel.teamAadObjectId });
        microsoftTeams.appInitialization.notifySuccess();
    };

    // token failure callback for SSO
    tokenFailedCallback = (error: string) => {
        console.error(`SSO failed: ${error}`);
        microsoftTeams.appInitialization.notifyFailure({
            reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
            message: this.props.t('errorAuthFailed'),
        });
    };

    // handler when a new agent is selected
    changeItems(members: TeamMember[]) {
        this.setState({ selectedAgents: members });
    }

    // handler when comment text area changes
    commentChanged: ComponentEventHandler<TextAreaProps> = (event, data) => {
        const newValue = (event.target as HTMLTextAreaElement).value;
        this.setState({ comment: newValue });
    };

    // handler when reassign button is clicked
    reassignClicked: ComponentEventHandler<ButtonProps> = async () => {
        this.setState({ saving: true });
        let consultDetails: ConsultDetails = null;
        try {
            consultDetails = await reassignConsult(
                this.state.token,
                this.props.match.params.requestId,
                this.state.selectedAgents,
                this.state.comment);
        } catch {
            console.log("something has gone wrong");
            return;
        }

        // close task module
        const taskModuleResult: TaskModuleResult = { type: 'consultDetailsResult', consultDetails };
        microsoftTeams.tasks.submitTask(taskModuleResult);
    };

    // renders the component
    render() {
        return (
            <div className="page" style={{ padding: "20px", display: "flex", flexDirection: "column", height: "100vh" }}>
                <Flex gap="gap.small" padding="padding.medium">
                    <div style={{ width: "100%" }}>
                        <Text content={this.props.t('staffLabel')} style={{ width: "100%" }} />
                        <TeamMemberPicker
                            teamAadObjectId={this.state.selectedTeamId}
                            appToken={this.state.token}
                            graphToken={this.state.graphToken}
                            onChange={this.changeItems.bind(this)}
                            onError={this.props.alertHandler.bind(this)} />
                    </div>
                </Flex>
                <Flex gap="gap.small" padding="padding.medium">
                    <div style={{ width: "100%" }}>
                        <Text content={this.props.t('noteLabel')} style={{ width: "100%" }} />
                        <TextArea
                            fluid
                            variables={{ height: '80px' }}
                            placeholder={this.props.t('notePlaceholder')}
                            value={this.state.comment}
                            onChange={this.commentChanged}
                        />
                    </div>
                </Flex>
                <Flex style={{ flex: 1 }}></Flex>
                <Divider />
                <Flex hAlign="end" gap="gap.small">
                    <Button
                        content={this.props.t('reassignButton')}
                        primary
                        style={{ marginLeft: "auto" }}
                        onClick={this.reassignClicked}
                        loading={this.state.saving}
                        disabled={this.state.saving || !this.state.canReassign} />
                </Flex>
            </div>
        );
    }
}

export default withTranslation(['consultReassignModal', 'common'])(ConsultReassignModal);