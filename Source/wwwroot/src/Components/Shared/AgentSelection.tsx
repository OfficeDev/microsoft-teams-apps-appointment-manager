import * as React from "react";
import { RadioGroup, RadioGroupItemProps, ShorthandValue, Avatar, ComponentEventHandler, Flex, Text, ChevronEndIcon, Layout } from "@fluentui/react-northstar";
import { withTranslation, WithTranslation } from 'react-i18next';

// local imports
import { PhotoUtil } from "../../Utils/PhotoUtil";
import { AgentAvailability } from "../../Models/AgentAvailability";

// component properties
export interface AgentSelectionProps extends WithTranslation {
    agents: AgentAvailability[];
    graphToken: string;
    photoUtil: PhotoUtil;
    onAgentSelectionChanged?: (agent: AgentAvailability) => void;
}

// component state
export interface AgentSelectionState {
    agents: AgentAvailability[];
    selectedIndex: number;
}

// AgentSelection component
class AgentSelection extends React.Component<AgentSelectionProps, AgentSelectionState> {
    constructor(props: AgentSelectionProps) {
        super(props);
        this.state = {
            agents: this.props.agents,
            selectedIndex: -1,
        };
    }

    // component did mount
    componentDidMount = () => {
        // process photos for each agent
        this.state.agents.forEach((agent: AgentAvailability, index: number) => {
            agent.photo = this.props.photoUtil.emptyPic;
            this.props.photoUtil.getGraphPhoto(this.props.graphToken, agent.id).then((uri: string) => {
                const items = this.state.agents;
                items[index].photo = uri;
                this.setState({ agents: items });
            });
        });
    };

    // handles selection change
    agentSelectionChanged: ComponentEventHandler<RadioGroupItemProps> = (_, data) => {
        const selectedIndex = (data.value as number);
        this.setState({ selectedIndex });
        if (this.props.onAgentSelectionChanged) {
            this.props.onAgentSelectionChanged(this.state.agents[selectedIndex]);
        }
    };

    // renders the component
    render() {
        const timeSlotButtons = this.state.agents.map((agent: AgentAvailability, index: number) => {
            return {
                key: index,
                label: (
                    <Flex style={{ marginLeft: "-30px" }}>
                        <Layout start={<Avatar image={agent.photo} name={agent.displayName} />}
                            main={
                                <Flex style={{ paddingLeft: "15px", paddingTop: "6px" }}>
                                    <Text content={agent.displayName} weight="bold"></Text>
                                    <Text content={`(${agent.timeBlocks.length} ${this.props.t("slotsAvailableLabel")})`} style={{ paddingLeft: "6px", display: (agent.timeBlocks.length === 0) ? "none" : "block" }}></Text>
                                </Flex>}
                            end={<ChevronEndIcon />}
                            style={{ width: "100%" }} />
                    </Flex>
                ),
                value: index,
                className: `boxed ${this.state.selectedIndex === index ? "selected" : "unselected"}`,
                styles: { marginBottom: "1em" },
                variables: {
                    indicatorBorderColorDefault: "transparent",
                    indicatorBackgroundColorChecked: "transparent",
                    indicatorBorderColorDefaultHover: "transparent",
                },
            } as ShorthandValue<RadioGroupItemProps>;
        });

        return (
            <>
                <Text content={this.props.t("agentsLabel")} size="small" className="tmSectionTitle" />
                <RadioGroup
                    vertical
                    items={timeSlotButtons}
                    onCheckedValueChange={this.agentSelectionChanged}
                />
            </>
        );
    }
}

export default withTranslation(["consultAssignModal", "common"])(AgentSelection);