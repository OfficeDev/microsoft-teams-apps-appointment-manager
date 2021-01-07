import { CheckboxProps, DropdownItemProps } from "@fluentui/react-northstar";

export interface CheckboxItem<T> extends CheckboxProps {
    data: T;
}

export interface DropdownItem<T> extends DropdownItemProps {
    key: string;
    data: T;
}