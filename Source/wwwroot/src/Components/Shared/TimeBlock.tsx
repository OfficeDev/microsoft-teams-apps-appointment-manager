import * as React from "react";
import { Text, Flex } from '@fluentui/react-northstar';
import { TimeBlock as TimeBlockModel } from "../../Models/TimeBlock";
import moment from "moment";

// component properties for the alert component
export interface TimeBlockProps {
    timeBlock: TimeBlockModel;
    style?: React.CSSProperties;
}

// TimeBlock component for displaying a datetime range
class TimeBlock extends React.Component<TimeBlockProps> {
    // renders the component
    render(): React.ReactNode {
        const startMoment = moment(this.props.timeBlock.startDateTime);
        const endMoment = moment(this.props.timeBlock.endDateTime);

        if (!startMoment.isValid() || !endMoment.isValid()) {
            return null;
        }

        return (
            <Flex style={this.props.style}>
                <Text content={startMoment.format('ll')} weight="bold"></Text>
                <Text content={`${startMoment.format('LT')} - ${endMoment.format('LT')}`} style={{ paddingLeft: "6px" }}></Text>
            </Flex>
        );
    }
}

export default TimeBlock;