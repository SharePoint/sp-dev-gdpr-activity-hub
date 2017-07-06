import * as React from 'react';
import styles from '../GdprHierarchy.module.scss';

import IHierarchyListItemProps from './IHierarchyListItemProps';
import IHierarchyListItemState from './IHierarchyListItemState';

import { GDPRUtility } from '../../../../components/GDPRUtility';

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

/**
 * Persona
 */
import {
  Persona,
  PersonaSize,
  PersonaPresence
} from 'office-ui-fabric-react/lib/Persona';

export default class HierarchyListItem extends React.Component<IHierarchyListItemProps, IHierarchyListItemState> {

    /**
     * Main constructor for the component
     */
    constructor(props: IHierarchyListItemProps) {
      super();
      
      this.state = {
        item: props.item,
      };
    }    

    public render(): JSX.Element {

      let siteUrl: string = this.props.context.pageContext.site.absoluteUrl;

      return (
        this.state.item ?
          <div className={ css("ms-Grid-row", styles.hierarchyItem) }>
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12">
              <Persona
                imageUrl={ GDPRUtility.getPersonaImage(siteUrl, this.state.item.loginName) }
                imageInitials={ GDPRUtility.getInitials(this.state.item.fullName) }
                primaryText={ this.state.item.fullName }
                secondaryText={ this.state.item.role }
                size={ PersonaSize.regular  }
                presence={ PersonaPresence.none }
                hidePersonaDetails={ false }
              />
            </div>
          </div>
        : null
      );
  }
}