import { RequestStatus } from "../Apis/ConsultApi";

export interface ConsultScheduleTabSettings {
    statuses: RequestStatus[],
    categories: string[],
}