import { IWebPartContext} from '@microsoft/sp-webpart-base';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

import { ISPTaxonomyPickerProps } from './ISPTaxonomyPickerProps';

/**
 * @interface
 * Generic Term Object (abstract interface)
 */
export interface ISPTermObject {
  name: string;
  guid: string;
}

/**
 * @class
 * Service implementation to manage term stores in SharePoint
 * Basic implementation taken from: https://oliviercc.github.io/sp-client-custom-fields/
 */
export class SPTermStoreService {

  private context: IWebPartContext;
  private props: ISPTaxonomyPickerProps;
  private taxonomySession: string;
  private formDigest: string;

  /**
   * @function
   * Service constructor
   */
  constructor(props: ISPTaxonomyPickerProps){
      this.props = props;
      this.context = props.context;
  }

  /**
   * @function
   * Gets the collection of term stores in the current SharePoint env
   */
  public getTermsFromTermSet(termSet: string): Promise<ISPTermObject[]> {
    if (Environment.type === EnvironmentType.SharePoint ||
        Environment.type === EnvironmentType.ClassicSharePoint) {
      //First gets the FORM DIGEST VALUE
      var contextInfoUrl: string = this.context.pageContext.web.absoluteUrl + "/_api/contextinfo";
      var httpPostOptions: ISPHttpClientOptions = {
        headers: {
          "accept": "application/json",
          "content-type": "application/json"
        }
      };
      return this.context.spHttpClient.post(contextInfoUrl, SPHttpClient.configurations.v1, httpPostOptions).then((response: SPHttpClientResponse) => {
        return response.json().then((jsonResponse: any) => {
          this.formDigest = jsonResponse.FormDigestValue;

          //Build the Client Service Request
          var clientServiceUrl = this.context.pageContext.web.absoluteUrl + '/_vti_bin/client.svc/ProcessQuery';
          var data = '<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="JavaScript Client" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectIdentityQuery Id="3" ObjectPathId="1" /><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><ObjectPath Id="8" ObjectPathId="7" /><Query Id="9" ObjectPathId="7"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="Terms" SelectAll="true"><Query SelectAllProperties="false"><Properties /></Query></Property></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="1" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="4" ParentId="1" Name="GetDefaultSiteCollectionTermStore" /><Method Id="7" ParentId="4" Name="GetTermSetsByName"><Parameters><Parameter Type="String">' + termSet + '</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>';
          httpPostOptions = {
            headers: {
              'accept': 'application/json',
              'content-type': 'application/json',
              "X-RequestDigest": this.formDigest
            },
            body: data
          };
          return this.context.spHttpClient.post(clientServiceUrl, SPHttpClient.configurations.v1, httpPostOptions).then((serviceResponse: SPHttpClientResponse) => {
            return serviceResponse.json().then((serviceJSONResponse: Array<any>) => {

                let result: Array<ISPTermObject> = new Array<ISPTermObject>();

                serviceJSONResponse.forEach((child: any) => {
                    if (child != null && child['_ObjectType_'] !== undefined)
                    {
                        var termSetCollectionType = child['_ObjectType_'];
                        if (termSetCollectionType === "SP.Taxonomy.TermSetCollection") 
                        {
                            var childTermSets = child['_Child_Items_'];
                            childTermSets.forEach((ts: any) => {
                                
                                var termSetType = ts['_ObjectType_'];
                                if (termSetType === "SP.Taxonomy.TermSet") 
                                {
                                    var termCollection = ts['Terms'];
                                    var childTerms = termCollection['_Child_Items_'];
                                    childTerms.forEach((t: any) => {
                                        var termType = t['_ObjectType_'];
                                        if (termType === "SP.Taxonomy.Term") 
                                        {
                                            result.push({ guid: this.cleanGuid(t['Id']), name: t["Name"] });
                                        }
                                    });
                                }
                            });
                        }
                    }
                });

                return(result);
            });
          });

        });
      });
    }
    else
    {
      return (new Promise<Array<ISPTermObject>>((resolve, reject) => {
        resolve(new Array<ISPTermObject>());
      }));
    }
  }

  /**
   * @function
   * Clean the Guid from the Web Service response
   * @param guid
   */
  private cleanGuid(guid: string): string {
    if (guid !== undefined)
      return guid.replace('/Guid(', '').replace('/', '').replace(')', '');
    else
      return '';
  }
}