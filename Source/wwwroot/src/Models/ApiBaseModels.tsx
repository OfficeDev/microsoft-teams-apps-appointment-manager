export interface BaseModel {
    id: string;
    createdDateTime: string;
}

export interface CreatedByUserBaseModel extends BaseModel {
    createdById: string;
    createdByName: string;
}

export interface IdName {
    id: string;
    displayName: string;
}