import * as React from "react";
import { RadioGroup, RadioGroupItemProps, ShorthandValue, ComponentEventHandler, Layout, ChevronEndIcon, Flex, Text } from '@fluentui/react-northstar';
import { withTranslation, WithTranslation } from 'react-i18next';

// local imports
import TimeBlock from "./TimeBlock";
import * as TBModel from "../../Models/TimeBlock";


export interface TimeBlockSelectionProps extends WithTranslation {
    timeBlocks: TBModel.TimeBlock[];
    onTimeBlockSelectionChanged: (timeBlock: TBModel.TimeBlock) => void;
}

export interface TimeBlockSelectionState {
    timeBlocks: TBModel.TimeBlock[];
    selectedIndex: number;
}

// TimeBlockSelection component
class TimeBlockSelection extends React.Component<TimeBlockSelectionProps, TimeBlockSelectionState> {
    constructor(props: TimeBlockSelectionProps) {
        super(props);
        this.state = {
            timeBlocks: [...this.props.timeBlocks].sort((a, b) => new Date(a.startDateTime).getTime() - new Date(b.startDateTime).getTime()),
            selectedIndex: -1,
        };
    }

    // event raised when an item is selected
    timeBlockSelectionChanged: ComponentEventHandler<RadioGroupItemProps> = (_, data) => {
        const selectedIndex = (data.value as number);
        this.setState({ selectedIndex });
        if (this.props.onTimeBlockSelectionChanged) {
            this.props.onTimeBlockSelectionChanged(this.state.timeBlocks[selectedIndex]);
        }
    };

    // renders the component
    render = () => {
        const timeSlotButtons = this.state.timeBlocks.map((timeSlot: TBModel.TimeBlock, index: number) => {
            return {
                key: timeSlot.startDateTime,
                label: (
                    <Flex style={{ marginLeft: "-30px" }}>
                        <Layout main={<TimeBlock timeBlock={timeSlot} />}
                            end={<ChevronEndIcon />}
                            style={{ width: "100%" }} />
                    </Flex>),
                name: "freeSlots",
                value: index,
                className: `boxed ${this.state.selectedIndex === index ? 'selected' : 'unselected'}`,
                styles: { marginBottom: '1em' },
                variables: {
                    indicatorBorderColorDefault: "transparent",
                    indicatorBackgroundColorChecked: "transparent",
                    indicatorBorderColorDefaultHover: "transparent",
                },
            } as ShorthandValue<RadioGroupItemProps>;
        });

        return (
            <>
                <Text content={this.props.t("timesLabel")} size="small" className="tmSectionTitle" />
                <RadioGroup
                    vertical
                    items={timeSlotButtons}
                    onCheckedValueChange={this.timeBlockSelectionChanged} />
            </>
        );
    };
}

export default withTranslation(["consultAssignModal", "common"])(TimeBlockSelection);