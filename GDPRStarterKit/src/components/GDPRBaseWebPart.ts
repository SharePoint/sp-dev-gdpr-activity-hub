import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';

import * as strings from 'gdprInsertRequestStrings';
import { IGDPRBaseWebPartProps } from './IGDPRBaseWebPartProps';

import pnp from "sp-pnp-js";

export abstract class GdprBaseWebPart extends BaseClientSideWebPart<IGDPRBaseWebPartProps> {

  private listsOptions: IPropertyPaneDropdownOption[];

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
      pnp.setup({
        spfxContext: this.context
      });      
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    this.fetchLists().then((response) => {
      this.listsOptions = response;
      this.context.propertyPane.refresh();
    });

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown('targetList', {
                  label: strings.TargetListFieldLabel,
                  options: this.listsOptions
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private fetchLists(): Promise<IPropertyPaneDropdownOption[]> {

    return pnp.sp.web.lists.filter("Hidden eq false").get().then((response) => {
        var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
        response.map((list: any) => {
            options.push( { key: list.Id, text: list.Title });
        });

        return options;
    });
  }
}
