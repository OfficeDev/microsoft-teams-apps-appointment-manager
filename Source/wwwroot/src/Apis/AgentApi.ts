export async function getTeamMembers(token: string, teamId: string): Promise<TeamMember[]> {
    return fetch("/api/agents", {
        method: "GET",
        headers: new Headers({
            "Authorization": "Bearer " + token,
            "teamAadObjectId": teamId,
        }),
    }).then(async (response) => {
        if (!response.ok) {
            const message = await response.json();
            throw new Error(`Unable to get the agent list: ${message['reason']}`);
        }
        return response.json();
    }).then((jsonResponse: TeamMember[]) => {
        return jsonResponse;
    });
}

export async function updateAgent(token: string, agentId: string, agent: Partial<Agent>): Promise<void> {
    return fetch(`/api/agent/${agentId}`, {
        method: "PATCH",
        headers: new Headers({
            "Authorization": "Bearer " + token,
            "Content-Type": "application/json",
        }),
        body: JSON.stringify(agent),
    }).then(response => {
        if (!response.ok) {
            throw new Error(`UpdateAgent API call failed with status ${response.status}`);
        }
    });
}

interface Agent {
    userPrincipalName: string;
    aadObjectId: string;
    teamsId: string;
    serviceUrl: string;
    bookingsStaffMemberId: string;
    locale: string;
}

export interface TeamMember {
    id: string;
    displayName: string;
}