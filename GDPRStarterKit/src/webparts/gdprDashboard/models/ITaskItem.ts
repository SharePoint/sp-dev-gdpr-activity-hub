interface ITaskItem {
  id: number;
  title: string;
  assigneeId: number;
  assigneeLoginName: string;
  assigneeFullName: string;
  dueDate: Date;
  completed: boolean;
}

export default ITaskItem;