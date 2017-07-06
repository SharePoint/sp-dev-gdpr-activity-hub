import {
  IWebPartContext
} from '@microsoft/sp-webpart-base';

export interface ISPLookupItemsPickerProps {
  context: IWebPartContext;
  sourceListId: string;
  label: string;
  placeholder: string;
  itemLimit?: number;
  required: boolean;

  onChanged?: (itemsIds: number[]) => void;
}
