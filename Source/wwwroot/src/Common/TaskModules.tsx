import { TaskInfo } from "@microsoft/teams-js";
import { TFunction } from "i18next";
import { ChannelMapping } from "../Apis/ChannelApi";
import { ConsultDetails } from "../Apis/ConsultApi";

export interface ConsultDetailsTaskModuleResult {
    type: 'consultDetailsResult';
    consultDetails: ConsultDetails;
}

export interface ChannelMappingTaskModuleResult {
    type: 'channelMappingResult';
    channelMapping: ChannelMapping;
}

export type TaskModuleResult = ConsultDetailsTaskModuleResult | ChannelMappingTaskModuleResult;

export function assignSelfTaskModule(requestId: string, t: TFunction): TaskInfo {
    return {
        url: `https://${window.location.host}/consult/assign/${requestId}/self`,
        title: t('tmTitleAssignSelf'),
        height: 600,
        width: 600,
    };
}

export function assignToOtherTaskModule(requestId: string, t: TFunction): TaskInfo {
    return {
        url: `https://${window.location.host}/consult/assign/${requestId}/other`,
        title: t('tmTitleAssignOther'),
        height: 600,
        width: 600,
    };
}

export function detailsTaskModule(requestId: string, t: TFunction): TaskInfo {
    return {
        url: `https://${window.location.host}/consult/detail/${requestId}`,
        title: t('tmTitleDetails'),
        height: 550,
        width: 600,
    };
}

export function reassignTaskModule(requestId: string, t: TFunction): TaskInfo {
    return {
        url: `https://${window.location.host}/consult/reassign/${requestId}`,
        title: t('tmTitleReassign'),
        height: 300,
        width: 600,
    };
}

export function addServiceTaskModule(t: TFunction): TaskInfo {
    return {
        url: `https://${window.location.host}/admin/service`,
        title: t('tmTitleAddService'),
        height: 550,
        width: 600,
    };
}

export function editServiceTaskModule(category: string, t: TFunction): TaskInfo {
    return {
        url: `https://${window.location.host}/admin/service/${category}`,
        title: t('tmTitleEditService'),
        height: 550,
        width: 600,
    };
}