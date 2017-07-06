import * as React from 'react';
import * as ReactDom from 'react-dom';

import GdprHierarchy from './components/GdprHierarchy';
import { IGdprHierarchyProps } from './components/IGdprHierarchyProps';
import { GdprBaseWebPart } from '../../components/GDPRBaseWebPart';

export default class GdprHierarchyWebPart extends GdprBaseWebPart {

  private _gdprDashboardComponent: GdprHierarchy;

  public render(): void {
    const element: React.ReactElement<IGdprHierarchyProps> = React.createElement(
      GdprHierarchy,
      {
        context: this.context,
        targetList: this.properties.targetList
      }    
    );

    this._gdprDashboardComponent = <GdprHierarchy>ReactDom.render(element, this.domElement);
  }
}
