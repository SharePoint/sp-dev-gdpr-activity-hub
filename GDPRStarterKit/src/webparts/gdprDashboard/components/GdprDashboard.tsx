import * as React from 'react';
import styles from './GdprDashboard.module.scss';
import { IGdprDashboardProps } from './IGdprDashboardProps';
import { IGdprDashboardState } from './IGdprDashboardState';

import ITaskItem from '../models/ITaskItem';
import TaskList from "./TaskList/TaskList";

import { Environment, EnvironmentType } from '@microsoft/sp-core-library';

/**
 * Common Infrastructure
 */
import {
  BaseComponent,
  assign,
  autobind
} from 'office-ui-fabric-react/lib/Utilities';

/**
 * Button
 */
import { PrimaryButton, DefaultButton, CommandButton, Button, IButtonProps } from 'office-ui-fabric-react/lib/Button';

import { default as pnp, PermissionKind } from "sp-pnp-js";

export default class GdprDashboard extends React.Component<IGdprDashboardProps, IGdprDashboardState> {

   /**
   * Main constructor for the component
   */
  constructor(props: IGdprDashboardProps) {
    super(props);
    
    this.state = {
      taskItems : [],
      currentUserIsAdmin: false,
      filterByCurrentUser: true,
    };

    this.readCurrentUserIsAdmin();
  }

  public componentWillReceiveProps(props: IGdprDashboardProps) {
  	this.refreshTasksList();
  }

  public componentDidMount() {
  	this.refreshTasksList();
  }

  public render(): React.ReactElement<IGdprDashboardProps> {

    return (
      <div className={styles.gdprDashboard}>
        <div className={styles.container}>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12">
              {
                (this.state.currentUserIsAdmin ?
                <div>
                  <div className="ms-Grid">
                      <div className="ms-Grid-row">
                          <div className="ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2">
                            <CommandButton
                              icon="ThumbnailView"
                              onClick={ this._showMyTasks }>
                              My Tasks
                            </CommandButton>
                          </div>
                          <div className="ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2">
                            <CommandButton
                              icon="TaskManager"
                              onClick={ this._showAllTasks }>
                              All Tasks
                            </CommandButton>
                          </div>
                          <div className="ms-Grid-col ms-u-sm8 ms-u-md8 ms-u-lg8"></div>
                      </div>
                  </div>
                </div>
                : null)
              }
              <TaskList 
                context={ this.props.context }
                taskItems={ this.state.taskItems }
                onChangeTaskItem={ this._onChangeTaskItem }
              />
            </div>
          </div>
        </div>
      </div>
    );
  }

  private readCurrentUserIsAdmin() {
    
    pnp.sp.web.getCurrentUserEffectivePermissions().then(perms => {
      if (pnp.sp.web.hasPermissions(perms, PermissionKind.ManageWeb)) {
       this.state.currentUserIsAdmin = true;
       this.setState(this.state); 
      }
    });

  }

  @autobind
  private _showAllTasks(){
    this.state.filterByCurrentUser = false;
    this.state.taskItems = [];
    this.setState(this.state); 
    this.refreshTasksList();
  }

  @autobind
  private _showMyTasks(){
    this.state.filterByCurrentUser = true;
    this.state.taskItems = [];
    this.setState(this.state); 
    this.refreshTasksList();
  }

  private refreshTasksList() {
    if (this.props.targetList) {
      this.fetchTasks().then((r) => {
        this.state.taskItems = r;
        this.setState(this.state);
      });
    }
  }

  private fetchTasks(): Promise<ITaskItem[]> {
    if (Environment.type === EnvironmentType.SharePoint ||
        Environment.type === EnvironmentType.ClassicSharePoint) {

      var listOfTasks = pnp.sp.web.lists.getById(this.props.targetList).items
        .select("ID", "Title", "AssignedTo/Id", "AssignedTo/Title", "AssignedTo/Name", "Checkmark", "DueDate")
        .expand("AssignedTo");
          
      if (this.state.filterByCurrentUser) {
        listOfTasks = listOfTasks.filter("AssignedToId eq " + this.props.context.pageContext.legacyPageContext.userId);
      }

      return(listOfTasks.get().then((response) => {
          var tasks: Array<ITaskItem> = new Array<ITaskItem>();
          response.map((item: any) => {
            tasks.push( { 
              id: item.ID,
              title: item.Title,
              dueDate: new Date(item.DueDate),
              assigneeId: item.AssignedTo.length > 0 ? item.AssignedTo[0].Id : 0,
              assigneeLoginName: item.AssignedTo.length > 0 ? item.AssignedTo[0].Name : null,
              assigneeFullName: item.AssignedTo.length > 0 ? item.AssignedTo[0].Title : null,
              completed: (item.Checkmark == 1),
            });
          });

        return tasks;
      }));
    }
    else {
      return(new Promise<ITaskItem[]>((resolve, reject) => {
        resolve([]);
      }));
    }
  }
  
  @autobind
  private _onChangeTaskItem(task: ITaskItem) {

    pnp.sp.web.lists.getById(this.props.targetList).items.getById(task.id).update(
      {
        Checkmark: task.completed ? "1" : "0",
        Status: task.completed ? "Completed" : "Not Started",
      }
    ).then(i => { console.log(i); });

    return;
  }
}
