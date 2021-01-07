import * as React from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { Flex, Header, Text, MenuButton, Button, Table, MoreIcon, EditIcon, TrashCanIcon, MenuItemProps, ShorthandCollection, TableRowProps } from '@fluentui/react-northstar';
import { addServiceTaskModule, editServiceTaskModule, TaskModuleResult } from "../../Common/TaskModules";
import { withTranslation, WithTranslation } from "react-i18next";
import { AlertHandler } from "../../Models/AlertHandler";
import { deleteChannelMapping, getChannelMappings, getChannels, Channel, ChannelMapping } from "../../Apis/ChannelApi";
import { IdName } from "../../Models/ApiBaseModels";

// component properties
export interface AdminTabProps extends WithTranslation {
    alertHandler: AlertHandler;
}

// component state
export interface AdminTabState {
    mappings: ChannelMapping[];
    channels: Channel[];
    token: string;
}

// AdminTab component
class AdminTab extends React.Component<AdminTabProps, AdminTabState> {
    constructor(props: AdminTabProps) {
        super(props);
        this.state = {
            mappings: [],
            channels: [],
            token: "",
        };

        // perform the SSO request
        const authTokenRequest = {
            successCallback: this.tokenCallback,
            failureCallback: (error: string) => { console.log("Failure: " + error); },
        };
        microsoftTeams.authentication.getAuthToken(authTokenRequest);
    }

    // token callback from getAuthToken
    tokenCallback = (token: string) => {
        this.setState({ token: token });
        const loaded = { mappingsLoaded: false, channelsLoaded: false };

        // fetch existing channel mappings
        getChannelMappings(token).then((jsonResponse: ChannelMapping[]) => {
            this.setState({ mappings: jsonResponse });
            loaded.mappingsLoaded = true;

            // check if all data has finished loading
            if (loaded.channelsLoaded) {
                microsoftTeams.appInitialization.notifyAppLoaded();
                microsoftTeams.appInitialization.notifySuccess();
            }
        });

        // fetch channels
        getChannels(token).then((jsonResponse: Channel[]) => {
            this.setState({ channels: jsonResponse });
            loaded.channelsLoaded = true;

            // check if all data has finished loading
            if (loaded.mappingsLoaded) {
                microsoftTeams.appInitialization.notifyAppLoaded();
                microsoftTeams.appInitialization.notifySuccess();
            }
        });
    };

    // gets the channel for a channel id
    getChannel = (id: string) => {
        return this.state.channels.find(c => c.channelId === id);
    };

    // gets the channel named based on channel id
    getChannelName = (id: string) => {
        const channel = this.getChannel(id);
        return (channel) ? `${channel.teamName} - ${channel.channelName}` : "";
    }

    // deletes an existing mapping
    deleteMapping = (index: number) => {
        // start the delete
        deleteChannelMapping(this.state.token, this.state.mappings[index].id).then(() => {
            this.setState(prevState => ({
                mappings: [...prevState.mappings.slice(0, index), ...prevState.mappings.slice(index + 1)],
            }));
        });
    };

    editService = (category: string) => {
        const editServiceTaskInfo = editServiceTaskModule(category, this.props.t);
        microsoftTeams.tasks.startTask(editServiceTaskInfo, (err: string, result: unknown) => {
            const tmResult = result as TaskModuleResult;
            if (!err && tmResult?.type === 'channelMappingResult') {
                // this is an edit...splice the item in mappings list
                this.setState((prevState) => {
                    const mappings = prevState.mappings;
                    for (let i = 0; i < mappings.length; i++) {
                        if (mappings[i].category === category) {
                            mappings.splice(i, 1, tmResult.channelMapping);
                        }
                    }
                    return { mappings };
                });
            }
        });
    };

    addService = () => {
        const addServiceTaskInfo = addServiceTaskModule(this.props.t);
        microsoftTeams.tasks.startTask(addServiceTaskInfo, (err: string, result: unknown) => {
            const tmResult = result as TaskModuleResult;
            if (!err && tmResult?.type === 'channelMappingResult') {
                // this was an add...add the item to mappings list
                this.setState((prevState) => ({
                    mappings: [...prevState.mappings, tmResult.channelMapping],
                }));
            }
        });
    };

    flattenSupervisors = (supervisors: IdName[]) => {
        if (supervisors && supervisors.length > 0) {
            let flattened = "";
            supervisors.forEach((supervisor) => {
                const parts = supervisor.displayName.split(' ');
                flattened += `${parts[0]} ${parts[1].substring(0, 1)}, `;
            });
            if (flattened.length > 2) {
                flattened = flattened.substring(0, flattened.length - 2);
            }
            return flattened;
        }

        return "";
    };

    menuItemClick = (index: number, _evt: unknown, ctrl: MenuItemProps) => {
        if (ctrl.index === 0) {
            // launch dialog to edit the mapping
            this.editService(this.state.mappings[index].category);
        }
        else {
            // delete the mapping
            this.deleteMapping(index);
        }
    };

    // renders the component
    render() {
        const header = {
            key: "header",
            items: [
                { content: <Text weight="regular" content={this.props.t('serviceColHeader')} />, key: "service" },
                { content: <Text weight="regular" content={this.props.t('channelColHeader')} />, key: "channel" },
                { content: <Text weight="regular" content={this.props.t('bookingsColHeader')} />, key: "bookings" },
                { content: <Text weight="regular" content={this.props.t('supervisorsColHeader')} />, key: "supervisors" },
                { content: <Text weight="regular" content="" />, key: "actions", style: { maxWidth: "60px" } },
            ],
        };

        const rows: ShorthandCollection<TableRowProps> = [];
        this.state.mappings.forEach((mapping, index) => {
            rows.push({
                key: index,
                items: [
                    {
                        content: <Text content={mapping.category}></Text>,
                    },
                    {
                        content: <Text content={this.getChannelName(mapping.channelId)}></Text>,
                    },
                    {
                        content: <Text content={((mapping.bookingsBusiness) ? mapping.bookingsBusiness.displayName : "") + "/" + ((mapping.bookingsService) ? mapping.bookingsService.displayName : "")}></Text>,
                    },
                    {
                        content: <Text content={this.flattenSupervisors(mapping.supervisors)}></Text>,
                    },
                    {
                        content: <MenuButton
                            className="ellipse"
                            trigger={<Button icon={<MoreIcon />} title={this.props.t('moreActionsButton')} />}
                            menu={[
                                {
                                    index: 0,
                                    content: (
                                        <Flex>
                                            <EditIcon />
                                            <Text style={{ paddingLeft: "10px" }} content={this.props.t('editServiceButton')}></Text>
                                        </Flex>
                                    ),
                                },
                                {
                                    index: 1,
                                    content: (
                                        <Flex>
                                            <TrashCanIcon />
                                            <Text style={{ paddingLeft: "10px" }} content={this.props.t('deleteServiceButton')}></Text>
                                        </Flex>
                                    ),
                                },
                            ]}
                            onMenuItemClick={this.menuItemClick.bind(this, index)}
                        ></MenuButton>, style: { maxWidth: "60px" },
                    },
                ],
            });
        });

        return (
            <div className="page">
                <Flex gap="gap.small">
                    <Header as="h2" content={this.props.t('adminTabHeader')} style={{ margin: "0px" }} />
                    <Flex.Item push>
                        <Button onClick={this.addService} content={`+ ${this.props.t('addServiceButton')}`} primary />
                    </Flex.Item>
                </Flex>
                <Table compact header={header} rows={rows} aria-label="Compact view static table" />
            </div>
        );
    }
}

export default withTranslation(['adminTab', 'common'])(AdminTab);