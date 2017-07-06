import {
  IWebPartContext
} from '@microsoft/sp-webpart-base';

export interface IGdprInsertRequestProps {
  context: IWebPartContext;
  targetList: string;
}
