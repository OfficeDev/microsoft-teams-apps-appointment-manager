// enumeration for blocked reason
export enum BlockedReason {
    NotBlocked = "NotBlocked",
    AlreadyAssigned = "AlreadyAssigned",
    NoAvailabilitySelf = "NoAvailabilitySelf",
    NoAvailabilityTeam = "NoAvailabilityTeam",
    NotAuthorized = "NotAuthorized",
    NotAuthorizedAndNoAvailability = "NotAuthorizedAndNoAvailability",
}

// enumeration for assignment actions
export enum AssignmentAction {
    Assign = "Assign",
    Back = "Back",
    Cancel = "Cancel",
    NoCancel = "NoCancel",
    Override = "Override"
}

// enumeration for assignment actions
export enum AssignmentStep {
    Blocked = 0,
    SelectAgent = 1,
    SelectTime = 2,
    Comments = 3
}