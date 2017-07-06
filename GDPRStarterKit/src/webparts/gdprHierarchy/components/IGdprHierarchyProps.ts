import {
  IWebPartContext
} from '@microsoft/sp-webpart-base';

export interface IGdprHierarchyProps {
  context: IWebPartContext;
  targetList: string;
}