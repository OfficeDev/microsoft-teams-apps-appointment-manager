import * as React from "react";
import { Flex, Button } from '@fluentui/react-northstar';
import { withTranslation, WithTranslation } from 'react-i18next';

// local imports
import { AssignmentAction } from "../../Models/AssignmentEnums";

export interface AssignmentActionsProps extends WithTranslation {
    actions: AssignmentAction[];
    actionClicked: (action: AssignmentAction) => void;
    saving: boolean;
}

// AssignmentActions component
class AssignmentActions extends React.Component<AssignmentActionsProps> {
    // fires when an action is clicked
    actionClicked = (action: AssignmentAction, _evt: unknown, _data: unknown) => {
        if (this.props.actionClicked) {
            this.props.actionClicked(action);
        }
    };

    // Gets the text for the action
    getActionText = (action: AssignmentAction): string => {
        switch (action) {
            case AssignmentAction.Assign:
                return this.props.t("assignButton");
            case AssignmentAction.Back:
                return this.props.t("backButton");
            case AssignmentAction.Cancel:
                return this.props.t("cancelButton");
            case AssignmentAction.NoCancel:
                return this.props.t("noCancelButton");
            case AssignmentAction.Override:
                return this.props.t("yesContinueButton");
        }
    };

    // renders the component
    render = () => {
        const btns = this.props.actions.map((action: AssignmentAction, index: number) => {
            return <Button
                content={this.getActionText(action)}
                onClick={this.actionClicked.bind(this, action)}
                primary={(action === AssignmentAction.Assign || action === AssignmentAction.Override)}
                loading={(this.props.saving && action === AssignmentAction.Assign)}
                disabled={(this.props.saving)}
            />;
        });
        return (
            <Flex hAlign="end" gap="gap.small" style={{ paddingBottom: "10px" }}>
                {btns}
            </Flex>
        );
    };
}

export default withTranslation(["consultAssignModal", "common"])(AssignmentActions);