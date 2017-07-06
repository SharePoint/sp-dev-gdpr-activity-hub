import * as React from 'react';
import styles from './GdprHierarchy.module.scss';

import IHierarchyItem from '../models/IHierarchyItem';
import { IGdprHierarchyProps } from './IGdprHierarchyProps';
import { IGdprHierarchyState } from './IGdprHierarchyState';

import pnp from "sp-pnp-js";

import { Environment, EnvironmentType } from '@microsoft/sp-core-library';

import { List } from 'office-ui-fabric-react';
import HierarchyListItem from './HierarchyListItem/HierarchyListItem';

/**
 * Common Infrastructure
 */
import {
  BaseComponent,
  assign,
  autobind,
  css
} from 'office-ui-fabric-react/lib/Utilities';

export default class GdprHierarchy extends React.Component<IGdprHierarchyProps, IGdprHierarchyState> {

  private roles = [
    { TermGuid: "8fb4774c-2f1a-4bea-b7a4-75ebfe74cbdc", Name: "Data Protection Officer", SortOrder: 1 },
    { TermGuid: "0cc951b5-612f-4d83-9e90-7993c3c0d15e", Name: "GDPR Controller", SortOrder: 2 },
    { TermGuid: "e28a2c76-eb00-4f5f-95e7-673b24b9b4dd", Name: "GDPR Processor", SortOrder: 3 },
  ];

  constructor(props: IGdprHierarchyProps) {
    super(props);

    this.state = {
      hierarchyItems: []
    };
  }

  public componentWillReceiveProps(props: IGdprHierarchyProps) {
    this._refreshHierarchy();
  }

  public componentDidMount() {
    this._refreshHierarchy();
  }

  public render(): React.ReactElement<IGdprHierarchyProps> {

    let dpoContacts: IHierarchyItem[] = this.state.hierarchyItems.filter(i => i.role == "Data Protection Officer");
    let controllerContacts: IHierarchyItem[] = this.state.hierarchyItems.filter(i => i.role == "GDPR Controller");
    let processorContacts: IHierarchyItem[] = this.state.hierarchyItems.filter(i => i.role == "GDPR Processor");

    return (
      <div className={styles.helloWorld}>
        <div className={ css('ms-Grid', styles.container) }>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12">
              <div className={ styles.GDPRHierarchyTitle }>GDPR Hierarchy</div>
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12">
              <List
                items={ dpoContacts }
                onRenderCell={ this._onRenderHierarchyItem }
                />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12">
              <List
                items={ controllerContacts }
                onRenderCell={ this._onRenderHierarchyItem }
                />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12">
              <List
                items={ processorContacts }
                onRenderCell={ this._onRenderHierarchyItem }
                />
            </div>
          </div>
        </div>
      </div>
    );
  }

  private _refreshHierarchy() {
    if (this.props.targetList) {
      this.fetchHierarchy().then((r) => {
        this.state.hierarchyItems = r;
        this.setState(this.state);
      });
    }
  }

  private fetchHierarchy(): Promise<IHierarchyItem[]> {
    if (Environment.type === EnvironmentType.SharePoint ||
        Environment.type === EnvironmentType.ClassicSharePoint) {

      return(pnp.sp.web.lists.getById(this.props.targetList).items
          .select("ID", "Title", "GDPRContactRole", "GDPRContactUser/Title", "GDPRContactUser/Name")
          .expand("GDPRContactUser")
          .get().then((response) => {
            var items: Array<IHierarchyItem> = new Array<IHierarchyItem>();
            response.map((item: any) => {
              items.push( { 
                id: item.ID,
                fullName: item.GDPRContactUser ? item.GDPRContactUser.Title : '',
                loginName: item.GDPRContactUser ? item.GDPRContactUser.Name : 0,
                role: this.roles.filter(r => r.TermGuid == item.GDPRContactRole.TermGuid)[0].Name,
                roleSortOrder: this.roles.filter(r => r.TermGuid == item.GDPRContactRole.TermGuid)[0].SortOrder,
                });
            });

        items = items.sort((i, t) => i.roleSortOrder - t.roleSortOrder);

        return items;
      }));
    }
    else {
      return(new Promise<IHierarchyItem[]>((resolve, reject) => {
        resolve([]);
      }));
    }
  }

  @autobind
  private _onRenderHierarchyItem(item: IHierarchyItem, index: number) {
    return (
      <HierarchyListItem 
        context={ this.props.context }
        item={ item }
        />
    );
  }   
}
