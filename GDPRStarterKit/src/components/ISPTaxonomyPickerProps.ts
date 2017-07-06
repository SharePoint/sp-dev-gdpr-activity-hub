import {
  IWebPartContext
} from '@microsoft/sp-webpart-base';

import { ISPTermObject } from './SPTermStoreService';

export interface ISPTaxonomyPickerProps {
  context: IWebPartContext;
  termSetName: string;
  label: string;
  placeholder: string;
  required: boolean;
  
  allowMultipleSelections?: boolean;
  excludeOfflineTermStores?: boolean;
  excludeSystemGroup?: boolean;
  displayOnlyTermSetsAvailableForTagging?: boolean;
  
  onChanged?: (terms: ISPTermObject[]) => void;
}
