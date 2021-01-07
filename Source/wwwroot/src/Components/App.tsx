import * as React from "react";
import { Route, Router } from "react-router-dom";
import * as microsoftTeams from "@microsoft/teams-js";
import { Provider, teamsTheme, teamsDarkTheme, teamsHighContrastTheme, Alert, ThemePrepared, Loader } from "@fluentui/react-northstar";
import { createBrowserHistory } from "history";
import moment from 'moment';
import { i18nInit } from './i18n';

// Shared Components
import RouteListener from "./Shared/RouteListener";
import Error from "./Views/Error";

// Tab Views
import AdminTab from "./Views/AdminTab";
import MyConsultsTab from "./Views/MyConsultsTab";
import ConsultScheduleTabConfig from "./Views/ConsultScheduleTabConfig";
import ConsultScheduleTab from "./Views/ConsultScheduleTab";

// Modal Views
import AddEditServiceModal from "./Views/AddEditServiceModal";
import ConsultAssignModal, { ConsultAssignModalProps } from "./Views/ConsultAssignModal";
import ConsultDetailModal from "./Views/ConsultDetailModal";
import ConsultReassignModal, { ConsultReassignModalProps } from "./Views/ConsultReassignModal";
import ConsultAttachModal from "./Views/ConsultAttachModal";
import { updateAgent } from "../Apis/AgentApi";
import { getSettings } from "../Apis/SettingsApi";
import { AlertHandler, AlertType } from "../Models/AlertHandler";

const browserHistory = createBrowserHistory();

type TeamsTheme = 'default' | 'dark' | 'contrast';

// component properties
export type AppProps = Record<string, never>;

// component state
export interface AppState {
    alert: string;
    alertType: AlertType;
    theme: TeamsTheme;
    initFinished: boolean;
}

// App component
export default class App extends React.Component<AppProps, AppState> {
    private teamsContext: microsoftTeams.Context;

    constructor(props: AppProps) {
        super(props);
        this.state = {
            alert: null,
            alertType: null,
            theme: "default",
            initFinished: false,
        };

        // initialize Microsoft Teams SDK...required on initial load
        microsoftTeams.initialize();

        // initialize handlers for theme changed and set initial theme
        microsoftTeams.getContext(async (teamsContext: microsoftTeams.Context) => {
            this.teamsContext = teamsContext;

            this.themeChanged(teamsContext.theme);

            const settings = await getSettings();

            const i18n = i18nInit(settings.defaultLocale);
            i18n.changeLanguage(teamsContext.locale);
            moment.locale([teamsContext.locale, settings.defaultLocale]);

            this.setState({ initFinished: true });
        });

        microsoftTeams.registerOnThemeChangeHandler(this.themeChanged);
    }

    componentDidMount(): void {
        // perform the SSO request
        const authTokenRequest: microsoftTeams.authentication.AuthTokenRequest = {
            successCallback: this.tokenCallback,
            failureCallback: (error) => console.error(`Failed to get token: ${error}`),
        };
        microsoftTeams.authentication.getAuthToken(authTokenRequest);
    }

    private tokenCallback = (token: string) => {
        // Asynchronously update agent's locale for later use
        updateAgent(token, this.teamsContext.userObjectId, {
            locale: this.teamsContext.locale,
        }).catch(error => {
            console.error(`Failed to update agent's locale: ${error}`);
        });
    };

    // handles Teams theme changes
    private themeChanged = (theme: string) => {
        this.setState({ theme: theme as TeamsTheme });
    };

    // fires when then route is changed
    private onRouteChanged = () => {
        // clear out the current error
        this.setState({ alert: null });
    };

    // toggles the alert message
    private onAlert: AlertHandler = (alert, alertType) => {
        this.setState({ alert: alert, alertType: alertType });
    };

    // dismisses the alert message
    private onAlertDismiss = () => {
        this.setState({ alert: null, alertType: null });
    };

    // renders the component
    render(): React.ReactNode {
        if (!this.state.initFinished) {
            return null;
        }

        // get router details
        const router = (
            <Router history={browserHistory}>
                {this.state.alert && <Alert content={this.state.alert.toString()} danger={this.state.alertType === "danger"} success={this.state.alertType === "success"} warning={this.state.alertType === "warning"} dismissible onVisibleChange={this.onAlertDismiss.bind(this)} />}
                <RouteListener routeChangedHandler={this.onRouteChanged.bind(this)}>
                    {/* Catch-all routes */}
                    <Route exact path="/" render={(props) => <Error {...props} />}/>
                    <Route exact path="/_=_" render={(props) => <Error {...props} />}/>

                    {/* Tab routes */}
                    <Route exact path="/admin" render={(props) => <AdminTab alertHandler={this.onAlert.bind(this)} {...props} />}/>
                    <Route path="/consult/my" render={(props) => <MyConsultsTab alertHandler={this.onAlert.bind(this)} {...props} />}/>
                    <Route path="/consult/scheduleconfig" render={(props) => <ConsultScheduleTabConfig alertHandler={this.onAlert.bind(this)} {...props} />}/>
                    <Route path="/consult/schedule/:config" render={(props) => <ConsultScheduleTab alertHandler={this.onAlert.bind(this)} {...props} />}/>

                    {/* Modal routes */}
                    <Route path="/admin/service/:category?" render={(props) => <AddEditServiceModal alertHandler={this.onAlert.bind(this)} {...props} />}/>
                    <Route path="/consult/assign/:requestId/:mode" render={(props: ConsultAssignModalProps) => <ConsultAssignModal alertHandler={this.onAlert.bind(this)} {...props} />}/>
                    <Route path="/consult/detail/:requestId" render={(props) => <ConsultDetailModal alertHandler={this.onAlert.bind(this)} {...props} />} />
                    <Route path="/consult/reassign/:requestId" render={(props: ConsultReassignModalProps) => <ConsultReassignModal alertHandler={this.onAlert.bind(this)} {...props} />} />
                    <Route path="/consult/attach/:conversationId/:files" render={(props) => <ConsultAttachModal alertHandler={this.onAlert.bind(this)} {...props}/>} />
                </RouteListener>
            </Router>
        );

        // get provider wrapper based on theme
        const fullDom = (
            <React.Suspense fallback={<Loader />}>
                <Provider theme={this.providerThemeMap[this.state.theme]}>
                    <div className={this.classThemeMap[this.state.theme]}>
                        {router}
                    </div>
                </Provider>
            </React.Suspense>
        );

        return (
            <div className="appWrapper">
                {fullDom}
            </div>
        );
    }

    private readonly providerThemeMap: Record<TeamsTheme, ThemePrepared> = {
        default: teamsTheme,
        dark: teamsDarkTheme,
        contrast: teamsHighContrastTheme,
    };

    private readonly classThemeMap: Record<TeamsTheme, string> = {
        default: 'vcDefaultTheme',
        dark: 'vcDarkTheme',
        contrast: 'vcContrastTheme',
    };
}