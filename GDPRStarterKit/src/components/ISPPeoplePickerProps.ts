import {
  IWebPartContext
} from '@microsoft/sp-webpart-base';

export interface ISPPeoplePickerProps {
  context: IWebPartContext;
  itemLimit?: number;
  label: string;
  placeholder: string;
  required?: boolean;

  onChanged?: (items: string[]) => void;
}
