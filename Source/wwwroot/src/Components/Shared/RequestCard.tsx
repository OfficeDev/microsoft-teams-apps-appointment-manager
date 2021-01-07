import * as React from "react";
import { Card, Layout, Avatar, Text, MenuButton, Button, Flex, Divider, MoreIcon, ParticipantRemoveIcon, AcceptIcon, RedoIcon, CalendarIcon, ParticipantAddIcon, PopupIcon, VideoCameraEmphasisIcon, ComponentEventHandler, MenuItemProps, ShorthandCollection } from '@fluentui/react-northstar';
import { ConsultDetails, RequestStatus } from "../../Apis/ConsultApi";
import { PhotoUtil } from "../../Utils/PhotoUtil";
import TimeBlock from "../Shared/TimeBlock";
import { Trans, withTranslation, WithTranslation } from 'react-i18next';

// component properties for the RequestCard component
export interface RequestCardProps extends WithTranslation {
    request: ConsultDetails;
    graphToken: string;
    meAadId: string;
    photoUtil: PhotoUtil;
    onAction?: (request: ConsultDetails, action: string) => void;
}

// component state
export interface RequestCardState {
    requestItemPhoto: string;
}

// RequestCard component for consult request
class RequestCard extends React.Component<RequestCardProps, RequestCardState> {
    constructor(props: RequestCardProps) {
        super(props);

        this.state = { requestItemPhoto: this.props.photoUtil.emptyPic };

        this.props.photoUtil.getGraphPhoto(this.props.graphToken, this.props.request.assignedToId).then((uri: string) => {
            this.setState({ requestItemPhoto: uri });
        });
    }

    itemClick: ComponentEventHandler<MenuItemProps & { action: string }> = (evt, ctrl) => {
        if (this.props.onAction) {
            this.props.onAction(this.props.request, ctrl.action);
        }
    };

    // renders the component
    render() {
        let status;
        const actions: ShorthandCollection<MenuItemProps> = [];
        switch (this.props.request.status) {
            case RequestStatus.Unassigned:
                status = (<Flex style={{ paddingTop: "5px", paddingBottom: "10px", color: "#B51C3B" }}><ParticipantRemoveIcon /><Text style={{ paddingLeft: "10px" }} content={this.props.t("stateUnassigned")} weight="bold" /></Flex>);
                actions.push({ content: (<Flex><ParticipantAddIcon /><Text style={{ paddingLeft: "10px" }} content={this.props.t("tmTitleAssignSelf")}></Text></Flex>), action: "assignMe" });
                actions.push({ content: (<Flex><ParticipantAddIcon /><Text style={{ paddingLeft: "10px" }} content={this.props.t("tmTitleAssignOther")}></Text></Flex>), action: "assignOther" });
                break;
            case RequestStatus.ReassignRequested:
                status = (<Flex style={{ paddingTop: "5px", paddingBottom: "10px", color: "#9F5811" }}><RedoIcon /><Text style={{ paddingLeft: "10px" }} content={this.props.t("stateReassign")} weight="bold" /></Flex>);
                actions.push({ content: (<Flex><PopupIcon /><Text style={{ paddingLeft: "10px" }} content={this.props.t("tmTitleDetails")}></Text></Flex>), action: "detail" });
                break;
            case RequestStatus.Assigned:
                // check if me or other so can be reused
                const isSelf = (this.props.request.assignedToId === this.props.meAadId);
                status = (
                    <Flex style={{ paddingTop: "5px", paddingBottom: "10px" }}>
                        <Avatar name={this.props.request.assignedToName} size="smallest" image={this.state.requestItemPhoto} />
                        <Trans i18nKey={isSelf ? "stateAssignedToYou" : "stateAssignedTo"} t={this.props.t} values={isSelf ? {} : { name: this.props.request.assignedToName }}>
                            <Text style={{ paddingLeft: "6px" }} />
                            <Text weight="bold" style={{ paddingLeft: "6px" }} />
                        </Trans>
                    </Flex>);
                actions.push({ content: (<Flex><ParticipantRemoveIcon /><Text style={{ paddingLeft: "10px" }} content={this.props.t("reassignButton")}></Text></Flex>), action: "requestReassign" });
                actions.push({ content: (<Flex><AcceptIcon /><Text style={{ paddingLeft: "10px" }} content={this.props.t("completeButton")}></Text></Flex>), action: "markComplete" });
                actions.push({ content: (<Flex><VideoCameraEmphasisIcon /><Text style={{ paddingLeft: "10px" }} content={this.props.t("joinButton")}></Text></Flex>), action: "join" });
                break;
            case RequestStatus.Completed:
                status = (<Flex style={{ paddingTop: "5px", paddingBottom: "10px", color: "#1F6C3D" }}><ParticipantRemoveIcon /><Text style={{ paddingLeft: "10px" }} content={this.props.t("stateCompleted")} weight="bold" /></Flex>);
                break;
        }

        const prefTimes = this.props.request.preferredTimes.map((pref) => {
            return (<TimeBlock timeBlock={pref} />);
        });

        return (
            <div style={{ paddingTop: "10px", paddingBottom: "10px" }}>
                <Card className="mobileCard">
                    <Card.Header>
                        <Layout gap="1rem"
                            start={<Avatar name={this.props.request.customerName} />}
                            main={<Text content={this.props.request.customerName} weight="bold" />}
                            end={<MenuButton
                                style={{ display: (actions.length === 0) ? "none" : "block" }}
                                trigger={<Button text iconOnly icon={<MoreIcon />} title={this.props.t("openMenuButton")} />}
                                onMenuItemClick={this.itemClick.bind(this)}
                                menu={actions}
                            />}
                        />
                    </Card.Header>
                    <Card.Body>
                        <Text content={this.props.request.query} />
                        <Flex style={{ display: (!this.props.request.assignedTimeBlock) ? "none" : "inline-block", paddingTop: "10px" }}>
                            <CalendarIcon />
                            <div style={{ display: "inline-block" }}>
                                {(this.props.request.status !== RequestStatus.Unassigned) ?
                                    (<TimeBlock style={{ paddingLeft: "10px" }} timeBlock={this.props.request.assignedTimeBlock} />)
                                    : (<></>)
                                }
                            </div>
                        </Flex>
                        {status}
                        {(this.props.request.status === "Unassigned") ? (
                            <>
                                <Divider />
                                <Text content={this.props.t("preferredTimesLabel")} />
                                {prefTimes}
                            </>
                        ) : (<></>)}
                    </Card.Body>
                </Card>
            </div>);
    }
}

export default withTranslation(["consultScheduleTab", "common"])(RequestCard);