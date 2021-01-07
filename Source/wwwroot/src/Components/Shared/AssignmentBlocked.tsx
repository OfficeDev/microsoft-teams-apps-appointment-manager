import * as React from "react";
import { Flex, Text, ExclamationTriangleIcon, Segment, CalendarIcon } from "@fluentui/react-northstar";
import moment from "moment";
import { withTranslation, WithTranslation } from 'react-i18next';

// local imports
import { BlockedReason } from "../../Models/AssignmentEnums";
import { MeetingDetail } from "../../Apis/ConsultApi";

// component properties
export interface AssignmentBlockedProps extends WithTranslation {
    reason: BlockedReason;
    meetingDetail?: MeetingDetail[];
}

// AssignmentBlocked component
class AssignmentBlocked extends React.Component<AssignmentBlockedProps> {
    // get text for the UI
    getText = () => {
        switch (this.props.reason) {
            case BlockedReason.NotAuthorizedAndNoAvailability:
                return {
                    line1: this.props.t("unauthorizedNoAvailHeader"),
                    line2: this.props.t("unauthorizedNoAvailText"),
                    line3: this.props.t("unauthorizedNoAvailQuestion"),
                };
            case BlockedReason.NotAuthorized:
                return {
                    line1: this.props.t("unauthorizedHeader"),
                    line2: this.props.t("unauthorizedText"),
                    line3: this.props.t("unauthorizedQuestion"),
                };
            case BlockedReason.NoAvailabilitySelf:
                return {
                    line1: this.props.t("noAvailableTimesHeader"),
                    line2: this.props.t("noAvailableTimesText"),
                    line3: this.props.t("noAvailableTimesQuestion"),
                };
            case BlockedReason.NoAvailabilityTeam:
                return {
                    line1: this.props.t("noAvailableTimesTeamHeader"),
                    line2: this.props.t("noAvailableTimesTeamText"),
                    line3: this.props.t("noAvailableTimesTeamQuestion"),
                };
        }
    };

    private getMeetingSection() {
        const meetingDets = this.props.meetingDetail ?? [];
        if (meetingDets.length === 0) {
            return null;
        }

        const firstMeeting = meetingDets[0];
        const subject = firstMeeting.subject;
        const startMoment = moment(firstMeeting.meetingTime.startDateTime);
        const endMoment = moment(firstMeeting.meetingTime.endDateTime);
        const meetingElement = (
            <Segment
                color="brand"
                inverted
                style={{
                    borderRadius: "4px",
                    padding: "0.5em 3em 0.5em 0.5em",
                    width: "100%",
                    marginBottom: "6px",
                }}>
                <Flex gap="gap.small" vAlign="center">
                    <CalendarIcon size="large" />
                    <Flex column>
                        <Text
                            content={subject}
                            weight="semibold"
                        />
                        <Text
                            content={`${startMoment.format("llll")} - ${endMoment.format("LT")}`}
                            size="small"
                        />
                    </Flex>
                </Flex>
            </Segment>
        );

        return (
            <div>
                {meetingElement}
            </div>
        );
    }

    // renders the component
    render() {
        const text = this.getText();
        const meetingDetails = (this.props.meetingDetail) ? this.getMeetingSection() : <br/>;
        return (
            <div className="emphasized" style={{ padding: "18px 32px 30px" }}>
                <br />
                <Flex hAlign="center">
                    <ExclamationTriangleIcon outline size="large" />
                </Flex><br />
                <Flex hAlign="center" >
                    <Text
                        content={text.line1}
                        size="large"
                        weight="semibold"
                    />
                </Flex><br />
                <Flex hAlign="center" >
                    <Text
                        content={text.line2}
                        size="small"
                        weight="regular"
                    />
                </Flex>
                <br />
                {meetingDetails}
                <br />
                <Flex space="evenly"
                    vAlign="center">
                    <Text content={text.line3} />
                </Flex>
            </div>
        );
    }
}

export default withTranslation(["consultAssignModal", "common"])(AssignmentBlocked);