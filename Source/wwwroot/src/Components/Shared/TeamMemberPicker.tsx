import * as React from "react";
import { Dropdown } from '@fluentui/react-northstar';

import { getTeamMembers, TeamMember } from '../../Apis/AgentApi';
import { PhotoUtil } from '../../Utils/PhotoUtil';
import { withTranslation, WithTranslation } from "react-i18next";
import { AlertHandler } from "../../Models/AlertHandler";
import { DropdownItem } from "../../Models/ComponentItems";

// component properties
export interface TeamMemberPickerProps extends WithTranslation {
    teamAadObjectId: string;
    appToken: string;
    graphToken: string;
    value?: TeamMember[];
    onChange?: (teamMembers: TeamMember[]) => void;
    placeholder?: string;
    onError?: AlertHandler;
}

// component state
export interface TeamMemberPickerState {
    teamMembers: DropdownItem<TeamMember>[];
    selectedTeamMembers: DropdownItem<TeamMember>[];
}

// TeamMemberPicker component
class TeamMemberPicker extends React.Component<TeamMemberPickerProps, TeamMemberPickerState> {
    constructor(props: TeamMemberPickerProps) {
        super(props);

        const value = props.value ?? [];
        this.state = {
            teamMembers: [],
            selectedTeamMembers: value.map(m => this.getTeamMemberDropdownItem(m)),
        };
    }

    photoUtil: PhotoUtil = new PhotoUtil();

    componentDidUpdate(prevProps: TeamMemberPickerProps, prevState: TeamMemberPickerState) {
        // look for teamAadObjectId change
        if (prevProps.teamAadObjectId !== this.props.teamAadObjectId) {
            // load the team members for this team (asynchronously)
            this.refreshMembers(this.props.teamAadObjectId);
            return;
        }

        // if selected members are being controlled, look for selected members change
        if (this.props.value && !this.areSelectedMembersSame(this.props.value)) {
            // update state with new selected team members
            const selectedMembers: DropdownItem<TeamMember>[] = this.props.value.map(m => this.getTeamMemberDropdownItem(m));
            this.updateSelectedMembers(selectedMembers);
        }
    }

    // refreshes the members list in state based on a selected team id provided in properties
    private refreshMembers = async (teamId: string) => {
        this.setState({ teamMembers: [] });

        let response: TeamMember[] = [];
        try {
            response = await getTeamMembers(this.props.appToken, teamId);
        } catch (err) {
            if (this.props.onError) {
                this.props.onError(this.props.t('errorCannotGetMembers'), "danger");
            }
            return;
        } finally {
            // remove selected team members who aren't in the new team
            const filteredSelected = this.state.selectedTeamMembers.filter(selectedTeamMember => response.some(teamMember => teamMember.id === selectedTeamMember.data.id));
            if (filteredSelected.length !== this.state.selectedTeamMembers.length) {
                this.selectionChanged(filteredSelected);
            }
        }

        // update state with new team members
        this.setState({ teamMembers: response.map(item => this.getTeamMemberDropdownItem(item)) });

        // lazy load photos for team members
        response.forEach(teamMember => {
            this.photoUtil.getGraphPhoto(this.props.graphToken, teamMember.id).then(photo => {
                this.setState(prevState => {
                    const modifiedMembers = [...prevState.teamMembers];
                    const index = modifiedMembers.findIndex(m => m.data.id === teamMember.id);
                    modifiedMembers[index].image = photo;
                    return {
                        teamMembers: modifiedMembers,
                    };
                });
            });
        });
    };

    // handles internal selection changes
    private selectionChanged(selectedItems: DropdownItem<TeamMember>[]) {
        // if selected members aren't being controlled, update selected members in state
        if (!this.props.value) {
            this.updateSelectedMembers(selectedItems);
        }

        // raises onChange event if being used
        if (this.props.onChange) {
            this.props.onChange(selectedItems.map(i => i.data));
        }
    }

    // update state with new selected team members
    private updateSelectedMembers(newSelectedMembers: DropdownItem<TeamMember>[]) {
        this.setState({ selectedTeamMembers: newSelectedMembers });

        // lazy load photos for new selected team members
        newSelectedMembers.forEach(newSelectedTeamMember => {
            this.photoUtil.getGraphPhoto(this.props.graphToken, newSelectedTeamMember.data.id).then(photo => {
                this.setState(prevState => {
                    const modifiedMembers = [...prevState.selectedTeamMembers];
                    const index = modifiedMembers.findIndex(m => m.data.id === newSelectedTeamMember.data.id);
                    modifiedMembers[index].image = photo;
                    return {
                        selectedTeamMembers: modifiedMembers,
                    };
                });
            });
        });
    }

    // checks if the given members list matches the currently selected team members
    private areSelectedMembersSame(members: TeamMember[]) {
        if (members.length !== this.state.selectedTeamMembers.length) {
            return false;
        }

        for (const member of members) {
            const match = this.state.selectedTeamMembers.find(s => s.data.id === member.id);
            if (!match) {
                return false;
            }
        }

        return true;
    }

    // creates an item for use in the dropdown
    private getTeamMemberDropdownItem(teamMember: TeamMember): DropdownItem<TeamMember> {
        return teamMember ? { key: teamMember.id, data: teamMember, header: teamMember.displayName, image: this.photoUtil.emptyPic } : null;
    }

    // renders the component
    render() {
        return (
            <Dropdown
                items={this.state.teamMembers}
                value={this.state.selectedTeamMembers}
                placeholder={this.props.placeholder}
                search
                multiple
                checkable
                fluid
                noResultsMessage={this.props.t('dropdownNoResults')}
                onChange={(_evt, ctrl) => this.selectionChanged(ctrl.value as DropdownItem<TeamMember>[])}
                disabled={this.state.teamMembers.length === 0}
            />
        );
    }
}

export default withTranslation('common')(TeamMemberPicker);