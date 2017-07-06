import ITaskItem from "../models/ITaskItem";

export interface IGdprDashboardState {
  taskItems: ITaskItem[];
  currentUserIsAdmin?: boolean;
  filterByCurrentUser?: boolean;
}