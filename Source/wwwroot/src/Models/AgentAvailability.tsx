import { TimeBlock } from "./TimeBlock";

export interface AgentAvailability {
    id: string;
    displayName: string;
    emailaddress: string;
    timeBlocks: TimeBlock[];
    photo: string;
}