import * as React from "react";
import { RouteComponentProps } from "react-router-dom";
import * as microsoftTeams from "@microsoft/teams-js";
import { Flex, Input, Text, Attachment, Divider, Button, ButtonProps, InputProps, ComponentEventHandler } from "@fluentui/react-northstar";
import { ExclamationTriangleIcon, FilesTxtIcon } from "@fluentui/react-icons-northstar";
import { addAttachmentToConsultRequest, ConsultAttachment, ConsultDetails, getConsultIdFromConversationID } from "../../Apis/ConsultApi";
import { withTranslation, WithTranslation } from "react-i18next";
import { AlertHandler } from "../../Models/AlertHandler";

// route parameters
type RouteParams = {
    // the ID of the conversation where the message is located
    conversationId: string;
    files: string;
}

// component properties
export interface ConsultAttachModalProps extends RouteComponentProps<RouteParams>, WithTranslation {
    alertHandler: AlertHandler;
}

// component state
export interface ConsultAttachModalState {
    token: string;
    attachments: string[];
    request: ConsultDetails;
    isAttachInProgress: boolean;
    titleMapping: Record<string, string>;
}

// ConsultAttachModal component
class ConsultAttachModal extends React.Component<ConsultAttachModalProps, ConsultAttachModalState> {
    constructor(props: ConsultAttachModalProps) {
        super(props);
        this.state = {
            token: null,
            attachments: null,
            request: null,
            isAttachInProgress: false,
            titleMapping: {},
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
        const attachmentsStr = atob(this.props.match.params.files);
        const attachmentsObj = JSON.parse(attachmentsStr);

        try {
            const response = await getConsultIdFromConversationID(token, this.props.match.params.conversationId);
            // save the consult request
            this.setState({ request: response });

        } catch (err) {
            console.error(`Getting consult request failed: ${err}`);
            this.props.alertHandler(`Error: ${err}`, 'danger');
        }

        // save token to state
        this.setState({ token: token });
        // save attachments list to state
        this.setState({ attachments: attachmentsObj });
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

    // extracts the filename from a path
    private getFilenameFromPath(path: string): string {
        return path.replace(/^.*[\\/]/, '');
    }

    // cancels the task module
    onCancel = () => {
        microsoftTeams.tasks.submitTask();
    };

    // handler when attach button is clicked
    attachClicked: ComponentEventHandler<ButtonProps> = async () => {
        this.setState({ isAttachInProgress: true });
        try {
            for (let i = 0; i < this.state.attachments.length; i++) {
                const path = this.state.attachments[i];
                let curTitle = this.state.titleMapping[path];
                if (curTitle === undefined) {
                    curTitle = '';
                }
                const attachmentToSend: Partial<ConsultAttachment> = {
                    uri: path,
                    filename: this.getFilenameFromPath(path),
                    title: curTitle,
                };
                await addAttachmentToConsultRequest(this.state.token, this.state.request.id, attachmentToSend);
            }
        } catch (err) {
            this.setState({ isAttachInProgress: false });
            const errorMessage = (err instanceof Error)
                ? err.message
                : this.props.t('errorCannotAttach');
            this.props.alertHandler(errorMessage, 'danger');
            return;
        }

        // close task module
        microsoftTeams.tasks.submitTask();
    };

    // handles input change
    inputChanged: ComponentEventHandler<InputProps & { value: string; }> = (_evt, ctrl) => {
        const mapping = this.state.titleMapping;

        const key = ctrl.variables['filename'];
        if (key !== undefined || key !== "") {
            mapping[key] = ctrl.value;
            this.setState({ titleMapping: mapping });
        }
    };

    // renders the footer of the second page for comments
    private renderPageFooter() {
        return (
            <>
                <Divider className="tmDivider" />
                <div className="tmButtonSet">
                    <Button
                        content={this.props.t('cancelButton')}
                        secondary
                        iconPosition="before"
                        disabled= {this.state.isAttachInProgress}
                        onClick={() => this.onCancel()}/>
                    <Button
                        content={this.props.t('attachButton')}
                        disabled={this.state.isAttachInProgress}
                        loading={this.state.isAttachInProgress}
                        onClick={this.attachClicked}
                        primary />
                </div>
            </>
        );
    }

    // renders attachments page
    private renderAttachmentsPage() {
        const taskInfo: microsoftTeams.TaskInfo = { height: 650, width: 600 };
        microsoftTeams.tasks.updateTask(taskInfo);

        const attachmentCases = this.state.attachments ? this.state.attachments : [];
        const attachmentElements = attachmentCases.map((attachmentInstance) => {
            return (
                <div>
                    <Flex hAlign="start" style={{ marginBottom: '-1em' }} >
                        <Input placeholder={this.props.t('addTitlePlaceholder')}
                            onChange={this.inputChanged.bind(this)}
                            fluid
                            variables={{ filename: attachmentInstance }}/>
                    </Flex><br />
                    <Flex style={{ paddingLeft: "1em" }} >
                        <Flex style={{ borderStyle: 'solid', borderColor: 'lightgray', borderWidth: 'thin', flex: 1 }}>
                            <Flex gap="gap.large" vAlign="center" style={{ padding: "0.5em", minWidth: "350px" }}>
                                <Attachment
                                    icon={<FilesTxtIcon />}
                                    header={this.getFilenameFromPath(attachmentInstance)} />
                            </Flex><br />
                        </Flex><br />
                    </Flex><br />


                </div>
            );
        });
        return (
            <div className="taskModule">
                <div className="tmBody">
                    <Flex hAlign="start">
                        <Text content={this.props.t('bookingHeader', { id: this.state.request.friendlyId })} size="medium" weight="bold"/>
                    </Flex><br />
                    <Flex hAlign="start">
                        <Text content={this.props.t('attachmentsLabel')} size="small" className="tmSectionTitle" style={{ marginBottom: '-1em' }} />
                    </Flex><br />
                    {this.state.attachments && attachmentElements}
                </div>
                <div className="footer">
                    {this.renderPageFooter()}
                </div>
            </div>
        );
    }

    // renders no attachment page
    private renderNoAttachmentsPage() {
        const taskInfo: microsoftTeams.TaskInfo = { height: 300, width: 600 };
        microsoftTeams.tasks.updateTask(taskInfo);

        return (
            <div className="taskModule">
                <div className="tmBody">
                    <div className="emphasized" style={{ padding: "38px 46px" }}>
                        <Flex column hAlign="center" gap="gap.small">
                            <ExclamationTriangleIcon outline size="larger" />
                            <Text
                                content={this.props.t('noAttachmentsHeader')}
                                size="large"
                                weight="semibold"
                            />
                            <Text
                                content={this.props.t('noAttachmentsText')}
                                weight="light"
                                align="center"
                            />
                        </Flex>
                    </div>
                </div>
            </div>
        );
    }

    // renders the component
    render() {
        const attachmentList = this.state.attachments;
        if (attachmentList === null) {
            return null;
        } else if (attachmentList.length === 0) {
            return (<>{this.renderNoAttachmentsPage()}</>);
        } else {
            return (<>{this.renderAttachmentsPage()}</>);
        }
    }
}

export default withTranslation(['consultAttachModal', 'common'])(ConsultAttachModal);