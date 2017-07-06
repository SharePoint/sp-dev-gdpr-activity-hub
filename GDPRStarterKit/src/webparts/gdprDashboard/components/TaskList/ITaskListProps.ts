import TaskOperationCallback from '../../models/TaskOperationCallback';

import ITaskItem from '../../models/ITaskItem';

import {
  IWebPartContext
} from '@microsoft/sp-webpart-base';

interface ITaskListProps {
    context: IWebPartContext;
    taskItems: ITaskItem[];
    onChangeTaskItem?: TaskOperationCallback;
}

export default ITaskListProps;