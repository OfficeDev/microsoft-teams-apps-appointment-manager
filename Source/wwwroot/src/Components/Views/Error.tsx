import * as React from "react";
import { withTranslation, WithTranslation } from "react-i18next";

// component properties
export type ErrorProps = WithTranslation;

// Error component
class Error extends React.Component<ErrorProps> {
    // renders the component
    render() {
        return (
            <div className="page">
                <h1>TODO: Error</h1>
            </div>
        );
    }
}

export default withTranslation()(Error);