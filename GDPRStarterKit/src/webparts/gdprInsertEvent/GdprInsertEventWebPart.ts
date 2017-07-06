import * as React from 'react';
import * as ReactDom from 'react-dom';

import GdprInsertEvent from './components/GdprInsertEvent';
import { IGdprInsertEventProps } from './components/IGdprInsertEventProps';
import { GdprBaseWebPart } from '../../components/GDPRBaseWebPart';

export default class GdprInsertEventWebPart extends GdprBaseWebPart {

  public render(): void {
    const element: React.ReactElement<IGdprInsertEventProps> = React.createElement(
      GdprInsertEvent,
      {
        context: this.context,
        targetList: this.properties.targetList,
      }
    );

    ReactDom.render(element, this.domElement);
  }
}
