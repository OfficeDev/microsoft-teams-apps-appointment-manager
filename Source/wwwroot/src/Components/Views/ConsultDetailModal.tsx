import * as React from "react";
import { Redirect, RouteComponentProps } from "react-router-dom";
import * as microsoftTeams from "@microsoft/teams-js";
import { Menu, Button, Flex, Text, Avatar, Dialog, ButtonProps, Divider, AppsIcon, CalendarIcon, Header, ParticipantRemoveIcon, QuestionCircleIcon, Attachment, MoreIcon, TextArea, FilesTxtIcon, MenuButton, DownloadIcon, ComponentEventHandler, TextAreaProps, ChevronDownMediumIcon, SyncIcon, AcceptIcon, Provider, MenuProps } from '@fluentui/react-northstar';

import { ConsultAssignModalLocationState } from "./ConsultAssignModal";
import { ActivityType, addNoteToConsult, completeConsult, ConsultActivity, ConsultDetails, getConsultDetails, RequestStatus } from '../../Apis/ConsultApi';
import { getGraphTokenUsingSsoToken } from "../../Apis/UtilApi";
import { PhotoUtil } from "../../Utils/PhotoUtil";
import TimeBlock from "../Shared/TimeBlock";

import { LocationDescriptor } from 'history';
import { withTranslation, WithTranslation } from "react-i18next";
import moment from "moment";
import { TaskModuleResult } from "../../Common/TaskModules";
import { AlertHandler } from "../../Models/AlertHandler";

// route parameters
type RouteParams = {
    requestId: string;
}

// component properties
export interface ConsultDetailModalProps extends RouteComponentProps<RouteParams>, WithTranslation {
    alertHandler: AlertHandler;
}

// component state
export interface ConsultDetailModalState {
    token: string;
    graphToken: string;
    request: ConsultDetails;
    detailsExpanded: boolean;
    activeTabIndex: number;
    notes: string;
    shouldRedirectTo: 'assignSelf' | 'assignOther' | 'reassign';
    isAddNoteInProgress: boolean;
    isCompleteInProgress: boolean;
    userImageUrls: { [id: string]: string };
    needCompleteConfirmation: boolean;
}

// ConsultDetailModal component
class ConsultDetailModal extends React.Component<ConsultDetailModalProps, ConsultDetailModalState> {
    constructor(props: ConsultDetailModalProps) {
        super(props);
        this.state = {
            token: null,
            graphToken: null,
            request: null,
            detailsExpanded: false,
            activeTabIndex: 0,
            notes: '',
            shouldRedirectTo: null,
            isAddNoteInProgress: false,
            isCompleteInProgress: false,
            userImageUrls: {},
            needCompleteConfirmation: false,
        };

        // perform the SSO request
        const authTokenRequest = {
            successCallback: this.tokenCallback,
            failureCallback: this.tokenFailedCallback,
        };
        microsoftTeams.authentication.getAuthToken(authTokenRequest);
    }

    // token failure callback for SSO
    tokenFailedCallback = (error: string) => {
        microsoftTeams.appInitialization.notifyFailure({
            reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
            message: this.props.t('errorAuthFailed'),
        });
    };

    // token callback from getAuthToken
    tokenCallback = async (token: string) => {
        // save token in state
        this.setState({ token });

        // start getting Graph token
        getGraphTokenUsingSsoToken(token).then(graphToken => {
            this.setState({ graphToken });
        });

        // get request from API and sort activities/notes in descending time order
        const consultDetails = await getConsultDetails(token, this.props.match.params.requestId);
        consultDetails.activities?.sort((a, b) => new Date(b.createdDateTime).getTime() - new Date(a.createdDateTime).getTime());
        consultDetails.notes?.sort((a, b) => new Date(b.createdDateTime).getTime() - new Date(a.createdDateTime).getTime());
        this.setState({ request: consultDetails });

        microsoftTeams.appInitialization.notifySuccess();
    };

    // assign to another agent
    assignToAnotherAgentClicked = () => {
        this.setState({ shouldRedirectTo: 'assignOther' });
    };

    // assign to me
    assignToMeClicked = () => {
        this.setState({ shouldRedirectTo: 'assignSelf' });
    };

    // request reassignment
    requestReassignClicked = () => {
        this.setState({ shouldRedirectTo: 'reassign' });
    };

    // marks the request complete
    markCompleteClicked = () => {
        this.setState({ needCompleteConfirmation: true });
    };

    // launches the online meeting
    joinCallClicked = () => {
        if (this.state.request.joinUri) {
            microsoftTeams.executeDeepLink(this.state.request.joinUri);
        }
        else {
            this.props.alertHandler(this.props.t('errorCannotJoin'), "warning");
        }
    };

    // toggles the tabbed menu
    menuChanged: ComponentEventHandler<MenuProps> = (evt, ctrl) => {
        if (typeof ctrl.activeIndex === 'number') {
            this.setState({ activeTabIndex: ctrl.activeIndex });
        }
    };

    // attachment menu clicked
    attachmentDownloadClicked = (attachmentUri: string) => {
        window.open(attachmentUri, "_blank");
    };

    // user typed in the notes text box
    notesChanged: ComponentEventHandler<TextAreaProps> = (event, data) => {
        const newValue = (event.target as HTMLTextAreaElement).value;
        this.setState({ notes: newValue });
    };

    // handler when cancel button is clicked
    cancelClicked: ComponentEventHandler<ButtonProps> = () => {
        // close task module
        this.setState({ needCompleteConfirmation: false });
    };

    // handler when complete button is clicked
    completeClicked: ComponentEventHandler<ButtonProps> = async () => {
        this.setState({ isCompleteInProgress: true });
        let consultDetails: ConsultDetails = null;
        try {
            consultDetails = await completeConsult(this.state.token, this.props.match.params.requestId);
        } catch {
            this.setState({ isCompleteInProgress: false });
            this.props.alertHandler(this.props.t('errorCannotComplete'), "warning");
        }
        finally {
            this.setState({ needCompleteConfirmation: false, isCompleteInProgress: false });
        }
        // close task module
        const taskModuleResult: TaskModuleResult = { type: 'consultDetailsResult', consultDetails };
        microsoftTeams.tasks.submitTask(taskModuleResult);
    };

    // add note
    addNoteClicked = async () => {
        this.setState({ isAddNoteInProgress: true });

        // call API to add note
        try {
            const createdNote = await addNoteToConsult(this.state.token, this.state.request.id, this.state.notes);

            // reset notes state
            // also manually add note to the request in state
            // this saves us from having to get the whole consult again
            this.setState(prevState => {
                const prevNotes = prevState.request.notes ?? [];
                return {
                    notes: '',
                    request: {
                        ...prevState.request,
                        notes: [createdNote, ...prevNotes],
                    },
                };
            });
        } catch {
            this.props.alertHandler(this.props.t('errorCannotAddNote'), 'warning');
        } finally {
            this.setState({ isAddNoteInProgress: false });
        }
    };

    // renders the component
    render() {
        const { request, shouldRedirectTo, activeTabIndex } = this.state;

        // in case render() is called before state gets updated
        if (!request) {
            return null;
        }

        // redirect to another task module
        if (shouldRedirectTo) {
            let pathname: string;
            if (shouldRedirectTo === 'assignSelf') {
                pathname = `/consult/assign/${request.id}/self`;
            } else if (shouldRedirectTo === 'assignOther') {
                pathname = `/consult/assign/${request.id}/other`;
            } else {
                pathname = `/consult/reassign/${request.id}`;
            }

            const toLocation: LocationDescriptor<ConsultAssignModalLocationState> = {
                pathname,
                state: { request },
            };
            return <Redirect to={toLocation} />;
        }

        const activeTabContent = activeTabIndex === 0
            ? this.renderOverviewTab()
            : (activeTabIndex === 1
                ? this.renderAttachmentsTab()
                : this.renderNotesTab());

        const tabMenuItems = [
            {
                key: "overview",
                content: this.props.t('overviewTab'),
            },
            {
                key: "attachments",
                content: this.props.t('attachmentsTab'),
            },
            {
                key: "internalNotes",
                content: this.props.t('notesTab'),
            },
        ];

        return (
            <>
                <div className="taskModule">
                    <div className="tmBody">
                        <Header as="h3" content={this.props.t('bookingHeader', { id: request.friendlyId })} className="tmHeading" style={{ marginBottom: '0.5em' }} />
                        <Flex gap="gap.smaller" vAlign="center" style={{ marginBottom: '0.2em' }}>
                            <AppsIcon outline />
                            <Text content={request.category} size="medium" weight="semibold" />
                        </Flex>
                        {request.status !== RequestStatus.Unassigned
                            && (
                                <Flex gap="gap.small" vAlign="center" style={{ marginBottom: '0.2em' }}>
                                    <CalendarIcon outline />
                                    <TimeBlock timeBlock={request.assignedTimeBlock}></TimeBlock>
                                </Flex>
                            )}
                        {this.renderActivityTimeline()}
                        <Menu defaultActiveIndex={0} items={tabMenuItems} underlined primary onActiveIndexChange={this.menuChanged.bind(this)} style={{ marginTop: '0.8em', marginBottom: '0.8em' }} />
                        {activeTabContent}
                    </div>
                    <div className="footer">
                        {(request.status === RequestStatus.Unassigned || request.status === RequestStatus.ReassignRequested) && this.renderUnassignedFooter()}
                        {request.status === RequestStatus.Assigned && this.renderAssignedFooter()}
                    </div>
                </div>
                <div className="page">
                    <Dialog
                        open={this.state.needCompleteConfirmation || this.state.isCompleteInProgress}
                        header={this.props.t('completeConfirm')}
                        confirmButton={{ content: this.props.t('confirmButton'), loading: this.state.isCompleteInProgress, disabled: this.state.isCompleteInProgress }}
                        cancelButton={{ content: this.props.t('cancelButton'), disabled: this.state.isCompleteInProgress }}
                        onCancel={this.cancelClicked}
                        onConfirm={this.completeClicked}
                    /></div></>
        );
    }

    // renders the activity timeline and the expand/collapse button
    renderActivityTimeline() {
        const { request, detailsExpanded } = this.state;
        // if request is unassigned, only show the "Unassigned" status
        if (request.status === RequestStatus.Unassigned) {
            return (
                <Flex gap="gap.smaller" vAlign="center">
                    <ParticipantRemoveIcon />
                    <Text content={this.props.t('stateUnassigned')} size="medium" color="red" weight="semibold" />
                </Flex>
            );
        }

        const detailsToggleButton = (
            <Button
                content={{
                    content: detailsExpanded ? this.props.t('collapseDetailsButton') : this.props.t('expandDetailsButton'),
                    size: 'small',
                }}
                text
                primary
                icon={{
                    content: <ChevronDownMediumIcon rotate={detailsExpanded ? 180 : 0} />,
                    styles: { marginLeft: '3px' },
                }}
                iconPosition="after"
                className="textOnly"
                onClick={() => this.setState(prevState => ({ detailsExpanded: !prevState.detailsExpanded }))} />
        );

        // if details are collapsed, only show the last activity
        if (!detailsExpanded) {
            return (
                <Flex gap="gap.smaller" vAlign="center">
                    {this.renderActivity(request.activities[0], false, false)}
                    {detailsToggleButton}
                </Flex>
            );
        }

        // details are expanded, so show all activities that can be rendered
        const validActivityTypes = [ActivityType.Assigned, ActivityType.ReassignRequested, ActivityType.Completed];
        const activities = request.activities
            .filter(activity => validActivityTypes.includes(activity.type));

        const activityElements = activities.map((activity, index, { length }) => {
            const isLastActivity = index === length - 1;
            return this.renderActivity(activity, true, !isLastActivity);
        });

        return (
            <>
                {activityElements}
                <div style={{ marginTop: '0.5em' }}>
                    {detailsToggleButton}
                </div>
            </>
        );
    }

    // renders a single activity in the activity timeline
    renderActivity(activity: ConsultActivity, isExpanded: boolean, showConnector: boolean) {
        const { type, createdDateTime, createdById, createdByName, activityForUserId, activityForUserName } = activity;

        const dateTimeString = moment(createdDateTime).format('L LT');

        let avatar: JSX.Element = null;
        let titleText: JSX.Element = null;
        let subtitleText: JSX.Element = null;
        if (type === ActivityType.Assigned) {
            avatar = <Avatar name={activityForUserName} image={this.getOrFetchUserPhoto(activityForUserId)} size="smallest" />;
            titleText =
                <Text
                    content={this.props.t('stateAssignedTo', { name: activityForUserName })}
                    size="medium"
                    weight="semibold" />;
            subtitleText =
                <Text
                    content={createdById === activityForUserId
                        ? this.props.t('assignedBySelf', { dateTime: dateTimeString })
                        : this.props.t('assignedBy', { name: createdByName, dateTime: dateTimeString })}
                    size="small"
                    weight="semilight" />;
        } else if (type === ActivityType.ReassignRequested) {
            avatar = <Avatar icon={<SyncIcon />} size="smallest" />;
            titleText =
                <Text
                    content={this.props.t('stateReassign')}
                    size="medium"
                    weight="semibold"
                    color="gold" />;
            subtitleText =
                <Text
                    content={this.props.t('requestedBy', { name: createdByName, dateTime: dateTimeString })}
                    size="small"
                    weight="semilight" />;
        } else if (type === ActivityType.Completed) {
            avatar = <AcceptIcon />;
            titleText =
                <Text
                    content={this.props.t('stateCompleted')}
                    size="medium"
                    weight="semibold" />;
            subtitleText =
                <Text
                    content={this.props.t('completedBy', { name: createdByName, dateTime: dateTimeString })}
                    size="small"
                    weight="semilight" />;
        }

        if (isExpanded) {
            return (
                <Flex gap="gap.smaller">
                    <Flex column>
                        {avatar}
                        {showConnector && <div className="vertLine" style={{ height: '21px' }}></div>}
                    </Flex>
                    <Flex column>
                        {titleText}
                        {subtitleText}
                    </Flex>
                </Flex>
            );
        } else {
            return (
                <Flex gap="gap.smaller" vAlign="center">
                    {avatar}
                    {titleText}
                </Flex>
            );
        }
    }

    renderOverviewTab() {
        const { customerName, customerPhone, customerEmail, query, status } = this.state.request;

        return (
            <>
                <Flex vAlign="center" gap="gap.smaller">
                    <Avatar name={customerName} />
                    <Flex column>
                        <Text content={customerName} weight="semibold" />
                        <Flex>
                            <Text content={customerPhone} />
                            <Divider vertical />
                            <Text content={customerEmail} />
                        </Flex>
                    </Flex>
                </Flex>

                <Provider.Consumer render={theme =>
                    <div style={{ padding: '1em', backgroundColor: theme.siteVariables.colorScheme.yellow.background1, marginTop: '0.8em' }}>
                        <Flex vAlign="center" gap="gap.smaller">
                            <QuestionCircleIcon size="small" variables={{ color: '#e4a512' }} />
                            <Text content={this.props.t('queryLabel')} weight="semibold" styles={{ color: theme.siteVariables.colorScheme.default.foregroundFocus }} />
                        </Flex>
                        <Text content={query} styles={{ color: theme.siteVariables.colorScheme.default.foregroundFocus }} />
                    </div>
                } />


                {status === RequestStatus.Unassigned ? this.renderPreferredTimeSection() : null}
            </>
        );
    }

    renderPreferredTimeSection() {
        const preferredTimes = this.state.request.preferredTimes ? [...this.state.request.preferredTimes] : [];
        preferredTimes.sort((a, b) => new Date(a.startDateTime).getTime() - new Date(b.startDateTime).getTime());

        const preferredTimeElements = preferredTimes.map((preferredTime) => {
            return (
                <TimeBlock timeBlock={preferredTime}></TimeBlock>
            );
        });

        return (
            <div style={{ marginTop: '0.8em' }}>
                <Text content={this.props.t('preferredTimeLabel')} size="small" className="tmSectionTitle" />
                {preferredTimeElements}
            </div>
        );
    }

    renderAttachmentsTab() {
        const attachments = this.state.request.attachments ?? [];
        if (attachments.length === 0) {
            return <Text content={this.props.t('noAttachments')} />;
        }

        return attachments.map(attachment =>
            <>
                <Text content={attachment.title} size="small" style={{ display: 'block', marginBottom: '0.5em' }} />
                <div style={{ display: 'block', marginBottom: '1.5em' }}>
                    <Attachment
                        icon={<FilesTxtIcon />}
                        header={attachment.filename}
                        action={{
                            content:
                                <MenuButton
                                    menu={[{ content: this.props.t('downloadButton'), icon: <DownloadIcon />, onClick: () => this.attachmentDownloadClicked(attachment.uri) }]}
                                    trigger={<Button text iconOnly icon={<MoreIcon />} />} />,
                        }} />
                </div>
            </>
        );
    }

    renderNotesTab() {
        const notes = this.state.request.notes ?? [];
        const noteElements = notes.map(note => {
            const createdDate = moment(note.createdDateTime).format('L LT');

            return (
                <div style={{ marginBottom: '0.8em' }}>
                    <Text content={note.text} />
                    <Flex vAlign="center" gap="gap.smaller">
                        <Avatar name={note.createdByName} image={this.getOrFetchUserPhoto(note.createdById)} size="smallest" />
                        <Text content={note.createdByName} size="small" />
                        <Text content={createdDate} size="small" />
                    </Flex>
                </div>
            );
        });
        return (
            <>
                <TextArea
                    placeholder={this.props.t('notesPlaceholder')}
                    fluid
                    variables={{ height: '100px' }}
                    value={this.state.notes}
                    onChange={this.notesChanged} />
                <Flex>
                    <Flex.Item push>
                        <Button
                            content={this.props.t('addNoteButton')}
                            text
                            primary
                            loading={this.state.isAddNoteInProgress}
                            disabled={this.state.notes.length === 0 || this.state.isAddNoteInProgress}
                            className="textOnly"
                            style={{ marginTop: '0.2em' }}
                            onClick={this.addNoteClicked} />
                    </Flex.Item>
                </Flex>

                {noteElements}
            </>
        );
    }

    renderUnassignedFooter() {
        return (
            <>
                <Divider className="tmDivider" />
                <div className="tmButtonSet">
                    <Button
                        content={this.props.t('assignMeButton')}
                        secondary
                        onClick={this.assignToMeClicked} />
                    <Button
                        content={this.props.t('assignOtherButton')}
                        secondary
                        onClick={this.assignToAnotherAgentClicked} />
                </div>
            </>
        );
    }

    renderAssignedFooter() {
        return (
            <>
                <Divider className="tmDivider" />
                <div className="tmButtonSet">
                    <Button
                        content={this.props.t('reassignButton')}
                        secondary
                        disabled={this.state.isCompleteInProgress}
                        onClick={this.requestReassignClicked} />
                    <Button
                        content={this.props.t('completeButton')}
                        secondary
                        disabled={this.state.isCompleteInProgress}
                        loading={this.state.isCompleteInProgress}
                        onClick={this.markCompleteClicked} />
                    <Button
                        content={this.props.t('joinButton')}
                        primary
                        disabled={this.state.isCompleteInProgress}
                        onClick={this.joinCallClicked} />
                </div>
            </>
        );
    }

    // returns object URL of user photo if previously fetched
    // otherwise, returns empty photo and starts fetching photo
    private getOrFetchUserPhoto(userId: string) {
        const { userImageUrls, graphToken } = this.state;

        if (userImageUrls[userId]) {
            return userImageUrls[userId];
        }

        // start fetching photo before returning
        if (graphToken) {
            this.photoUtil.getGraphPhoto(graphToken, userId).then(photoUrl => {
                this.setState(prevState => ({
                    userImageUrls: {
                        ...prevState.userImageUrls,
                        [userId]: photoUrl,
                    },
                }));
            });
        }

        // return empty photo for now
        return this.photoUtil.emptyPic;
    }

    private photoUtil: PhotoUtil = new PhotoUtil();
}

export default withTranslation(['consultDetailModal', 'common'])(ConsultDetailModal);