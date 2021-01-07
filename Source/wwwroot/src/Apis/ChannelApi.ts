import { BaseModel, IdName } from "../Models/ApiBaseModels";

export async function getChannels(token: string): Promise<Channel[]> {
    return fetch("/api/channel/", {
        method: "GET",
        headers: new Headers({
            "Authorization": "Bearer " + token,
        }),
    }).then(response => response.json() as Promise<Channel[]>);
}

export async function getChannelMappings(token: string): Promise<ChannelMapping[]> {
    return fetch("/api/channel/channelmappings", {
        method: "GET",
        headers: new Headers({
            "Authorization": "Bearer " + token,
        }),
    }).then(response => response.json() as Promise<ChannelMapping[]>);
}

export async function getChannelMappingsForTeam(token: string, teamId: string): Promise<ChannelMapping[]> {
    return fetch(`/api/channel/channelmappings/${teamId}`, {
        method: "GET",
        headers: new Headers({
            "Authorization": "Bearer " + token,
        }),
    }).then(response => response.json() as Promise<ChannelMapping[]>);
}

export async function patchChannelMapping(token: string, mapping: Partial<ChannelMapping>): Promise<boolean> {
    return fetch(`/api/channel/channelmappings/${mapping.id}`, {
        method: "PATCH",
        headers: new Headers({
            "Authorization": "Bearer " + token,
            "Accept": "application/json",
            "Content-Type": "application/json",
        }),
        body: JSON.stringify(mapping),
    }).then(response => response.ok);
}

export async function postChannelMapping(token: string, mapping: Partial<ChannelMapping>): Promise<ChannelMapping> {
    return fetch("/api/channel/channelmappings", {
        method: "POST",
        headers: new Headers({
            "Authorization": "Bearer " + token,
            "Accept": "application/json",
            "Content-Type": "application/json",
        }),
        body: JSON.stringify(mapping),
    }).then(response => response.json() as Promise<ChannelMapping>);
}

export async function deleteChannelMapping(token: string, mappingId: string): Promise<boolean> {
    return fetch(`/api/channel/channelmappings/${mappingId}`, {
        method: "DELETE",
        headers: new Headers({
            "Authorization": "Bearer " + token,
            "Accept": "application/json",
            "Content-Type": "application/json",
        }),
    }).then(response => response.ok);
}

export interface Channel extends BaseModel {
    tenantId: string;
    serviceUrl: string;
    teamId: string;
    teamAadObjectId: string;
    teamName: string;
    channelId: string;
    channelName: string;
}

export interface ChannelMapping extends BaseModel {
    channelId: string;
    category: string;
    bookingsBusiness: IdName;
    bookingsService: IdName;
    supervisors: IdName[];
}