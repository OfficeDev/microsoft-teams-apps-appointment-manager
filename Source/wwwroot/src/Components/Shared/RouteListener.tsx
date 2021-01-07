import * as React from "react";
import { RouteComponentProps, withRouter } from "react-router-dom";

// properties for the RouteListener component
export interface RouteListenerProps extends RouteComponentProps {
    routeChangedHandler: () => void;
}

// RouteListener component
class RouteListener extends React.Component<RouteListenerProps> {
    // componentWillMount
    componentWillMount() {
        this.props.history.listen((_location, _action) => {
            this.props.routeChangedHandler();
        });
    }

    // renders the component
    render() {
        return (
            <div className="container page-inner">{this.props.children}</div>
        );
    }
}

export default withRouter(RouteListener);