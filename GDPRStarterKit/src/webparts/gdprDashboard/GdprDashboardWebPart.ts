import * as React from 'react';
import * as ReactDom from 'react-dom';

import GdprDashboard from './components/GdprDashboard';
import { IGdprDashboardProps } from './components/IGdprDashboardProps';
import { GdprBaseWebPart } from '../../components/GDPRBaseWebPart';

export default class GdprDashboardWebPart extends GdprBaseWebPart {

  private _gdprDashboardComponent: GdprDashboard;

  public render(): void {
    const element: React.ReactElement<IGdprDashboardProps> = React.createElement(
      GdprDashboard,
      {
        context: this.context,
        targetList: this.properties.targetList,
      }    
    );

    this._gdprDashboardComponent = <GdprDashboard>ReactDom.render(element, this.domElement);
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    /*
    Check the property path to see which property pane feld changed. If the property path matches the dropdown, then we set that list
    as the selected list for the web part. 
    */
    if (propertyPath === 'targetList') {
      this._gdprDashboardComponent.props.targetList = this.properties.targetList;
    }

    /*
    Finally, tell property pane to re-render the web part. 
    This is valid for reactive property pane. 
    */
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }
}
