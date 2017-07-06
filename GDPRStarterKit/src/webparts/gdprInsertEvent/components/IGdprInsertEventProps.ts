import {
  IWebPartContext
} from '@microsoft/sp-webpart-base';

export interface IGdprInsertEventProps {
  context: IWebPartContext;
  targetList: string;
}
