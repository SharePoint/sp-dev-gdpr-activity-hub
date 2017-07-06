import * as React from 'react';
import styles from './GDPRStyles.module.scss';

import pnp from "sp-pnp-js";

import { ISPLookupItemsPickerProps } from './ISPLookupItemsPickerProps';
import { ISPLookupItemsPickerState } from './ISPLookupItemsPickerState';

/**
 * Common Infrastructure
 */
import {
  BaseComponent,
  assign,
  autobind,
  css
} from 'office-ui-fabric-react/lib/Utilities';

/**
 * Label
 */
import { Label } from 'office-ui-fabric-react/lib/Label';

import {
  IBasePickerProps,
  BasePickerListBelow,
  BaseAutoFill,
  IPickerItemProps
} from 'office-ui-fabric-react/lib/Pickers';

import { Icon } from 'office-ui-fabric-react/lib/Icon';

export interface ISPItemProps {
  itemId?: number;
  title?: string;
}

export interface ISPLookupItemPickerProps extends IBasePickerProps<ISPItemProps> {
}

export class SPLookupItemsPickerControl extends BasePickerListBelow<ISPItemProps, ISPLookupItemPickerProps> {  
}

export const SPLookupSuggestedItem: (itemProps: ISPItemProps) => JSX.Element = (itemProps: ISPItemProps) => {
  return (
    <div className={ styles.pickerRoot }>
      <span className={ styles.pickerSuggestedItem }>
        <span className={ styles.pickerSuggestedItemIcon }><i className="ms-Icon ms-Icon--QuickNote" aria-hidden="true"></i></span>
        <span className={ styles.pickerSuggestedItemText }>{ itemProps.title }</span>
      </span>
    </div>
  );
};

export const SPLookupSelectedItem: (itemProps: IPickerItemProps<ISPItemProps>) => JSX.Element = (itemProps: IPickerItemProps<ISPItemProps>) => {

  return (
    <div
      className={ css(styles.pickerRoot, styles.pickerSelectedItem) }
      key={ itemProps.item.itemId }
      data-selection-index={ itemProps.item.itemId }
      data-is-focusable={ true }>
      <span className={ styles.pickerSelectedItemIcon }>
        <i className="ms-Icon ms-Icon--QuickNote" aria-hidden="true"></i>
      </span>      
      <span className={ css('ms-TagItem-text', styles.pickerSelectedItemText) }>{ itemProps.item.title }</span>
      <span className={ css('ms-TagItem-close', styles.pickerSelectedItemClose) } onClick={ itemProps.onRemoveItem }>
        <i className="ms-Icon ms-Icon--Cancel" aria-hidden="true"></i>
      </span>
    </div>
  );
};

export class SPLookupItemsPicker extends React.Component<ISPLookupItemsPickerProps, ISPLookupItemsPickerState> {

  /**
   *
   */
  constructor(props: ISPLookupItemsPickerProps) {
    super(props);
    
    this.state = {
      itemsIds: [],
    };
  }

  public render(): React.ReactElement<ISPLookupItemsPickerProps> {

    pnp.setup({
      spfxContext: this.props.context,
    });      

    return (
      <div className={ css('ms-TextField', {'is-required': this.props.required }) }>
        <Label>{ this.props.label }</Label>
        <SPLookupItemsPickerControl
          onChange={ this._onChangeLookupItemsPicker }
          onResolveSuggestions={ this._onFilterChangedLookupItemsPicker }
          onRenderSuggestionsItem={ SPLookupSuggestedItem }
          onRenderItem={ SPLookupSelectedItem }
          getTextFromItem={ (props: ISPItemProps) => props.title }
          pickerSuggestionsProps={
            {
              suggestionsHeaderText: 'Suggested Items',
              noResultsFoundText: 'No Items Found',
              loadingText: 'Loading',
            }
          }
          />
      </div>
    );
  }

  @autobind
  private _onChangeLookupItemsPicker(items?: ISPItemProps[]): void{     
    
    /** Empty the array */
    this.state.itemsIds = new Array<number>();

    /** Fill it with new items */
    items.forEach((i: ISPItemProps) => {
        this.state.itemsIds.push(i.itemId);
    });
    this.setState(this.state);

    if (this.props.onChanged != null)
    {
        this.props.onChanged(this.state.itemsIds);
    }
  }

  @autobind
  private _onFilterChangedLookupItemsPicker(filterText: string, currentItems: ISPItemProps[]) : Promise<ISPItemProps[]> {
    
    if (filterText.length >= 3 && this.props.sourceListId) {

      let filteredLookupItems: ISPItemProps[] = new Array<ISPItemProps>();

      return(pnp.sp.web.lists.getById(this.props.sourceListId).items
        .filter("startswith(Title, '" + filterText + "')")
        .get().then((response) => {
          let items: Array<ISPItemProps> = new Array<ISPItemProps>();
          response.map((i: any) => {
              items.push( { itemId: i.Id, title: i.Title });
          });
          return items;
        }));
    }
  }
}


