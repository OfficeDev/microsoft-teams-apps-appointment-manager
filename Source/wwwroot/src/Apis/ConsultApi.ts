import { AgentAvailability } from "../Models/AgentAvailability";
import { BaseModel, CreatedByUserBaseModel, IdName } from "../Models/ApiBaseModels";
import { TimeBlock } from "../Models/TimeBlock";
import { Channel } from "./ChannelApi";

export async function getConsultDetails(token: string, consultId: string): Promise<ConsultDetails> {
    return fetch(`/api/request/${consultId}`, {
        method: "GET",
        headers: new Headers({
            "Authorization": "Bearer " + token,
        }),
    }).then(response => response.json() as Promise<ConsultDetails>);
}

export async function getConsultIdFromConversationID(token: string, conversationId: string): Promise<ConsultDetails> {
    return fetch(`/api/request/lookup/${encodeURIComponent(conversationId)}`, {
        method: "GET",
        headers: new Headers({
            "Authorization": "Bearer " + token,
        }),
    }).then(async (response) => {

        if (!response.ok) {
            const message = await response.json();
            throw new Error(`Unable to get the consult request: ${message['reason']}`);
        }

        return response.json() as Promise<ConsultDetails>;
    });
}

export async function getPostedChannel(token: string, reqId: string): Promise<Channel> {
    return fetch(`/api/request/${encodeURIComponent(reqId)}/channel`, {
        method: "GET",
        headers: new Headers({
            "Authorization": "Bearer " + token,
        }),
    }).then(response => response.json() as Promise<Channel>);
}

export async function getFilteredRequests(token: string, filter: RequestFilter): Promise<ConsultDetails[]> {
    return fetch("/api/request/filtered", {
        method: "POST",
        headers: new Headers({
            "Authorization": "Bearer " + token,
            "Accept": "application/json",
            "Content-Type": "application/json",
        }),
        body: JSON.stringify(filter),
    }).then(response => response.json() as Promise<ConsultDetails[]>);
}

export async function getMyRequests(token: string): Promise<ConsultDetails[]> {
    return fetch("/api/request", {
        method: "Get",
        headers: new Headers({
            "Authorization": "Bearer " + token,
        }),
    }).then((response) => response.json() as Promise<ConsultDetails[]>);
}

export async function getAvailability(token: string, timeConstraints: TimeBlock[], teamAadObjectId: string = null): Promise<AgentAvailability[]> {
    const uri = (teamAadObjectId) ? `/api/availability/${teamAadObjectId}` : "/api/availability";
    return fetch(uri, {
        method: "POST",
        headers: {
            "Authorization": "Bearer " + token,
            "Accept": "application/json",
            "Content-Type": "application/json",
        },
        body: JSON.stringify(timeConstraints),
    }).then(async (response) => {

        if (!response.ok) {
            const message = await response.json();
            console.log(message['reason']);
            return null;
        }

        return response.json() as Promise<AgentAvailability[]>;
    });
}

export async function reassignConsult(token: string, consultId: string, agents: IdName[], comments: string): Promise<ConsultDetails> {
    return fetch(`/api/request/${consultId}/reassign`, {
        method: "POST",
        headers: new Headers({
            "Authorization": "Bearer " + token,
            "Content-Type": "application/json",
        }),
        body: JSON.stringify({
            agents: agents,
            comments: comments,
        }),
    }).then(response => {
        if (!response.ok) {
            throw new Error(`ReassignConsult API call failed with status ${response.status}`);
        }

        return response.json() as Promise<ConsultDetails>;
    });
}

export async function getMeetingDetails(token: string, timeConstraints: TimeBlock): Promise<MeetingDetail[]> {
    return fetch(`/api/meetingDetails/`, {
        method: "POST",
        headers: {
            "Authorization": "Bearer " + token,
            "Accept": "application/json",
            "Content-Type": "application/json",
        },
        body: JSON.stringify(timeConstraints),
    }).then(async (response) => {

        if (!response.ok) {
            const message = await response.json();
            console.log(message['reason']);
            throw new Error(`Unable to get the meeting details: ${message['reason']}`);
        }

        return response.json() as Promise<MeetingDetail[]>;
    });
}

export async function assignConsult(token: string, consultId: string, timeBlock: TimeBlock, comments: string, agent: AgentAvailability): Promise<ConsultDetails> {
    return fetch(`/api/request/${consultId}/assign`, {
        method: "POST",
        headers: new Headers({
            "Authorization": "Bearer " + token,
            "Content-Type": "application/json",
        }),
        body: JSON.stringify({
            selectedTimeBlock: timeBlock,
            comments: comments,
            agent: agent,
        }),
    }).then(response => {
        if (!response.ok) {
            if (response.status === 400) {
                throw new Error(`The consult could not be assigned. Check with your Microsoft Bookings administrator to ensure you have proper permissions to schedule Bookings appointments.`);
            } else if (response.status === 403) {
                throw new Error(`The consult could not be assigned. It may already be assigned to someone.`);
            } else if (response.status === 404) {
                throw new Error(`The consult could not be found.`);
            } else {
                throw new Error(`The consult could not be assigned.`);
            }
        }

        return response.json() as Promise<ConsultDetails>;
    });
}

export async function addAttachmentToConsultRequest(token: string, consultId: string, attachment: Partial<ConsultAttachment>): Promise<void> {
    return fetch(`/api/request/attachment/${consultId}`, {
        method: "POST",
        headers: new Headers({
            "Authorization": "Bearer " + token,
            "Content-Type": "application/json",
        }),
        body: JSON.stringify(attachment),
    }).then(response => {
        if (!response.ok) {
            if (response.status === 400) {
                throw new Error(`Attachment or consult request not provided.`);
            } else if (response.status === 404) {
                throw new Error(`The consult could not be found.`);
            } else {
                throw new Error(`The consult could not be assigned.`);
            }
        }
    });
}


export async function completeConsult(token: string, consultId: string): Promise<ConsultDetails> {
    return fetch(`/api/request/${consultId}/complete`, {
        method: "POST",
        headers: new Headers({
            "Authorization": "Bearer " + token,
            "Content-Type": "application/json",
        }),
    }).then(response => {
        if (!response.ok) {
            throw new Error(`CompleteConsult API call failed with status ${response.status}`);
        }

        return response.json() as Promise<ConsultDetails>;
    });
}

export async function addNoteToConsult(token: string, consultId: string, notes: string): Promise<ConsultNote> {
    return fetch(`/api/request/${consultId}/notes`, {
        method: "POST",
        headers: new Headers({
            "Authorization": "Bearer " + token,
            "Content-Type": "application/json",
        }),
        body: JSON.stringify({
            text: notes,
        }),
    }).then(response => {
        if (!response.ok) {
            throw new Error(`AddNotes API call failed with status ${response.status}`);
        }

        return response.json() as Promise<ConsultNote>;
    });
}

export async function isSupervisor(token: string, consultId: string): Promise<boolean> {
    return fetch(`/api/request/${consultId}/issupervisor`, {
        method: "GET",
        headers: new Headers({
            "Authorization": "Bearer " + token,
        }),
    }).then(response => {
        if (!response.ok) {
            throw new Error(`isSupervisor API call failed with status ${response.status}`);
        }

        return response.json() as Promise<boolean>;
    });
}

export const enum RequestStatus {
    Unassigned = 'Unassigned',
    Assigned = 'Assigned',
    ReassignRequested = 'ReassignRequested',
    Completed = 'Completed',
}

export interface ConsultDetails extends BaseModel {
    customerName: string;
    customerPhone: string;
    customerEmail: string;
    query: string;
    preferredTimes: TimeBlock[];
    friendlyId: string;
    category: string;
    status: RequestStatus;
    assignedToId: string;
    assignedToName: string;
    assignedTimeBlock: TimeBlock;
    bookingsBusinessId: string;
    bookingsServiceId: string;
    bookingsAppointmentId: string;
    joinUri: string;
    activities: ConsultActivity[];
    notes: ConsultNote[]
    attachments: ConsultAttachment[];
}

export const enum ActivityType {
    Assigned = 'Assigned',
    ReassignRequested = 'ReassignRequested',
    Completed = 'Completed',
}

export interface ConsultActivity extends CreatedByUserBaseModel {
    type: ActivityType;
    activityForUserId: string;
    activityForUserName: string;
    comment: string;
}

interface ConsultNote extends CreatedByUserBaseModel {
    text: string;
}

export interface ConsultAttachment extends CreatedByUserBaseModel {
    id: string;
    filename: string;
    uri: string;
    title: string;
}

export interface MeetingDetail {
    subject: string;
    meetingTime: TimeBlock;
}

interface RequestFilter {
    categories: string[];
    statuses: RequestStatus[];
}