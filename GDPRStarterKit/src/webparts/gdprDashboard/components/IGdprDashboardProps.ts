import {
  IWebPartContext
} from '@microsoft/sp-webpart-base';

export interface IGdprDashboardProps {
  context: IWebPartContext;
  targetList: string;
}
