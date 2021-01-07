export type AlertHandler = (alert: string, alertType: AlertType) => void;
export type AlertType = 'danger' | 'warning' | 'success';