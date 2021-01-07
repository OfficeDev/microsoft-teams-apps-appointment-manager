import * as React from "react";
import { Text, TextArea, TextAreaProps, ComponentEventHandler } from '@fluentui/react-northstar';
import { withTranslation, WithTranslation } from 'react-i18next';

export interface AssignmentCommentsProps extends WithTranslation {
    comment: string;
    commentChanged: (comment: string) => void;
}

// AssignmentComments component
class AssignmentComments extends React.Component<AssignmentCommentsProps> {
    // fires when comments changed
    commentChanged: ComponentEventHandler<TextAreaProps> = (event, data) => {
        const newValue = (event.target as HTMLTextAreaElement).value;
        this.props.commentChanged(newValue);
    };

    // renders the component
    render = () => {
        return (
            <div style={{ width: "100%", paddingTop: "20px" }}>
                <Text content={this.props.t("notesLabel")} size="small" className="tmSectionTitle" />
                <TextArea
                    fluid
                    variables={{ height: '150px' }}
                    placeholder={this.props.t("typeHerePlaceholder")}
                    value={this.props.comment}
                    onChange={this.commentChanged}
                />
            </div>
        );
    };
}

export default withTranslation(["consultAssignModal", "common"])(AssignmentComments);