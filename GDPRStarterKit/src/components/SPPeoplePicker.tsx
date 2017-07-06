import * as React from 'react';

import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';

import { ISPPeoplePickerProps } from './ISPPeoplePickerProps';
import { ISPPeoplePickerState } from './ISPPeoplePickerState';
import { GDPRUtility } from './GDPRUtility';
import * as PeopleSearch from './PeopleSearchQuery';

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
 * People Picker
 */
import { IPersonaProps, PersonaPresence } from 'office-ui-fabric-react/lib/Persona';
import {
  IBasePickerSuggestionsProps,
  NormalPeoplePicker
} from 'office-ui-fabric-react/lib/Pickers';
import { IPersonaWithMenu } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';

/**
 * Label
 */
import { Label } from 'office-ui-fabric-react/lib/Label';

export interface IPeoplePickerExampleState {
  currentPicker?: number | string;
  delayResults?: boolean;
}

const suggestionProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: 'Suggested People',
  noResultsFoundText: 'No results found',
  loadingText: 'Loading'
};

export class SPPeoplePicker extends React.Component<ISPPeoplePickerProps, ISPPeoplePickerState> {

    /**
     * Constructor
     */
    constructor(props: ISPPeoplePickerProps) {
        super(props);
        
        this.state = {
            items: [],
        };
    }

  public render(): React.ReactElement<ISPPeoplePickerProps> {

    return (
      <div className={ css('ms-TextField', {'is-required': this.props.required }) }>
        <Label>{ this.props.label }</Label>
        <NormalPeoplePicker
            onChange={ this._onChangePeoplePicker }
            onResolveSuggestions={ this._onFilterChangedPeoplePicker }
            getTextFromItem={ (persona: IPersonaProps) => persona.primaryText }                
            pickerSuggestionsProps={ suggestionProps }
            className={ 'ms-PeoplePicker' }
            key={ 'normal' }
            />
      </div>
    );


  }

  @autobind
  private _onChangePeoplePicker(items?: IPersonaProps[]): void{     
    
    /** Empty the array */
    this.state.items = new Array<string>();

    /** Fill it with new items */
    items.forEach((i: IPersonaProps) => {
        this.state.items.push(i.secondaryText);
    });
    this.setState(this.state);

    if (this.props.onChanged != null)
    {
        this.props.onChanged(this.state.items);
    }
  }

  @autobind
  private _onFilterChangedPeoplePicker(filterText: string, currentPersonas: IPersonaProps[], limitResults?: number) {
    
    if (filterText) {

        let filteredPersonas: IPersonaProps[] = new Array<IPersonaProps>();

        let siteUrl: string = this.props.context.pageContext.site.absoluteUrl;
        let tenantBaseUrl: string = siteUrl.substring(0, siteUrl.indexOf("sharepoint.com") + 14);
        let peopleSearchUrl = tenantBaseUrl + "/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser";
        let imageBaseUrl =  tenantBaseUrl + "/_layouts/15/userphoto.aspx?size=S&accountname=";

        let searchRequest : ISPHttpClientOptions =  {
            body: JSON.stringify(new PeopleSearch.PeopleSearchQuery(
                {
                    QueryString: filterText,
                    MaximumEntitySuggestions: 10,
                    AllowEmailAddresses: true,
                    AllowOnlyEmailAddresses: false,
                    PrincipalType: 1,
                    PrincipalSource: 15,
                    SharePointGroupID: 0,
                })
            )
        };

        return(this.props.context.spHttpClient.post(
            peopleSearchUrl, 
            SPHttpClient.configurations.v1, 
            searchRequest)
                .then((response: SPHttpClientResponse) => {
                    return(response.json());
                })
                .then((people: PeopleSearch.IPeopleSearchQueryResult) => {
                    JSON.parse(people.value).forEach(p  => {
                        filteredPersonas.push({ 
                            primaryText: p.DisplayText,
                            secondaryText: p.Description,
                            imageInitials: GDPRUtility.getInitials(p.DisplayText), 
                            presence: PersonaPresence.none,
                            imageUrl: imageBaseUrl + p.Description,                            
                        });
                    });

                    return (filteredPersonas);
                }));
    } else {
      return [];
    }
  }

  private getInitials(fullname: string): string {
    var parts = fullname.split(' ');
    
    var initials = "";
    parts.forEach(p => {
        if (p.length > 0)
        {
            initials = initials.concat(p.substring(0, 1).toUpperCase());
        }
    });

    return (initials);
  }
}
