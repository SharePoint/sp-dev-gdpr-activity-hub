import { ISPTermObject } from './SPTermStoreService';

export interface ISPTaxonomyPickerState {
    terms: ISPTermObject[];
    loaded: boolean;
}