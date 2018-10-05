// State
import IListViewData from "./IListViewData";
export default interface IAppState {
    listData: IListViewData[];
    selectedItem?: { key: string | number | undefined; text: string };
    selectedFieldItem?: { key: string | number | undefined; text: string };
    selectedItems: string[];
    listViewColumns: any[];
    listViewItems: IListViewData[];
    disableView: boolean;
    selectListItem?: any;
    totalListItems: number;
    currentListPage: number;
    totalListPages: number;
    lastPageLoaded: number;
    showPanel: boolean;
  }