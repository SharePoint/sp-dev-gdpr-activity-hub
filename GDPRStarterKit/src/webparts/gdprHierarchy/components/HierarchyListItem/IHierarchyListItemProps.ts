import IHierarchyItem from '../../models/IHierarchyItem';

import {
  IWebPartContext
} from '@microsoft/sp-webpart-base';

interface IHierarchyListItemProps {
    context: IWebPartContext;
    item: IHierarchyItem;
}

export default IHierarchyListItemProps;