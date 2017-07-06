import * as React from 'react';
import * as strings from 'gdprDashboardStrings';

import ITaskListProps from './ITaskListProps';
import ITaskListState from './ITaskListState';

import ITaskItem from '../../models/ITaskItem';
import TaskListItem from '../TaskListItem/TaskListItem';

import { List } from 'office-ui-fabric-react';

/**
 * Common Infrastructure
 */
import {
  BaseComponent,
  assign,
  autobind
} from 'office-ui-fabric-react/lib/Utilities';

export default class TaskList extends React.Component<ITaskListProps, ITaskListState> {

  constructor(props: ITaskListProps) {
    super(props);

    this.state = {
      taskItems: []
    };
  }

  public componentWillReceiveProps(props: ITaskListProps) {
      if (props && props.taskItems) {
        this.setState({ taskItems: props.taskItems });
      }
  }

  public componentDidMount() {
    if (this.props && this.props.taskItems) {
      this.setState({ taskItems: this.props.taskItems });
    }
  }

  public render(): JSX.Element {

    // let tempItems : ITaskItem[] = [
    //   { id: 1, title: "Sample Task #01", assigneeId: 1, assigneeFullName: "Paolo Pialorsi", assigneeLoginName: "login", completed: true, dueDate: new Date() },
    //   { id: 2, title: "Sample Task #02", assigneeId: 2, assigneeFullName: "Paolo Pialorsi", assigneeLoginName: "login", completed: false, dueDate: new Date() },
    //   { id: 3, title: "Sample Task #03", assigneeId: 3, assigneeFullName: "Paolo Pialorsi", assigneeLoginName: "login", completed: false, dueDate: new Date() },
    //   { id: 4, title: "Sample Task #04", assigneeId: 4, assigneeFullName: "Paolo Pialorsi", assigneeLoginName: "login", completed: true, dueDate: new Date() },
    //   { id: 5, title: "Sample Task #05", assigneeId: 5, assigneeFullName: "Paolo Pialorsi", assigneeLoginName: "login", completed: false, dueDate: new Date() },
    // ];

    return (
        <div className="ms-Grid">
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2">{ strings.TaskCompletedColumnTitle }</div>
                <div className="ms-Grid-col ms-u-sm4 ms-u-md4 ms-u-lg4">{ strings.TaskAssigneeColumnTitle }</div>
                <div className="ms-Grid-col ms-u-sm4 ms-u-md4 ms-u-lg4">{ strings.TaskTitleColumnTitle }</div>
                <div className="ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2">{ strings.TaskDueDateColumnTitle }</div>
            </div>
            <List
              items={ this.state.taskItems }
              onRenderCell={ this._onRenderTaskItem }
              />
        </div>
    );
  }

  @autobind
  private _onRenderTaskItem(item: ITaskItem, index: number) {
    return (
      <TaskListItem 
        context={ this.props.context }
        task={ item }
        onChangeTaskItem={ this.props.onChangeTaskItem }
        />
    );
  }    
}