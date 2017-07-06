import * as React from 'react';

import ITaskListItemProps from './ITaskListItemProps';
import ITaskListItemState from './ITaskListItemState';

import { GDPRUtility } from '../../../../components/GDPRUtility';

/**
 * Common Infrastructure
 */
import {
  BaseComponent,
  assign,
  autobind
} from 'office-ui-fabric-react/lib/Utilities';

/**
 * Toggle
 */
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';

/**
 * Label
 */
import { Label } from 'office-ui-fabric-react/lib/Label';

/**
 * Persona
 */
import {
  Persona,
  PersonaSize,
  PersonaPresence
} from 'office-ui-fabric-react/lib/Persona';

export default class TaskListItem extends React.Component<ITaskListItemProps, ITaskListItemState> {

    /**
     * Main constructor for the component
     */
    constructor(props: ITaskListItemProps) {
      super();
      
      this.state = {
        task: props.task,
      };
    }    

    public render(): JSX.Element {

      let siteUrl: string = this.props.context.pageContext.site.absoluteUrl;

      return (
        this.state.task ?
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2">
              <Toggle
                onText=" "
                offText=" "
                checked={ this.state.task.completed }
                onChanged={ this._onChangedTask }
                />
            </div>
            <div className="ms-Grid-col ms-u-sm4 ms-u-md4 ms-u-lg4">
              <Persona
                imageUrl={ GDPRUtility.getPersonaImage(siteUrl, this.state.task.assigneeLoginName) }
                imageInitials={ GDPRUtility.getInitials(this.state.task.assigneeFullName) }
                primaryText={ this.state.task.assigneeFullName }
                size={ PersonaSize.extraExtraSmall  }
                presence={ PersonaPresence.none }
                hidePersonaDetails={ false }
              />
            </div>
            <div className="ms-Grid-col ms-u-sm4 ms-u-md4 ms-u-lg4">
              <Label>{ this.state.task.title }</Label>
            </div>
            <div className="ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2">
              <Label>{ this.state.task.dueDate.toLocaleDateString() }</Label>
            </div>
          </div>
        : null
      );
  }

  @autobind
  private _onChangedTask(newValue: boolean): void {
    this.state.task.completed = newValue;
    this.setState(this.state);

    if (this.props.onChangeTaskItem) {
      this.props.onChangeTaskItem(this.state.task);
    }
  }
}