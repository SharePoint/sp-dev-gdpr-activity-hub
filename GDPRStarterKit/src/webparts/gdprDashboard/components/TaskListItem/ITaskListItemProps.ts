import ITaskItem from '../../models/ITaskItem';
import TaskOperationCallback from '../../models/TaskOperationCallback';

import {
  IWebPartContext
} from '@microsoft/sp-webpart-base';

interface ITaskListItemProps {
  context: IWebPartContext;
  task: ITaskItem;
  onChangeTaskItem: TaskOperationCallback;
}

export default ITaskListItemProps;