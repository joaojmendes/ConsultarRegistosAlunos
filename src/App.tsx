// ***
// Poolyfills for old browsers IE9, IE10 e IE11
import "core-js/es6/map";
import "core-js/es6/set";
import "raf/polyfill";
import "core-js/es6/promise";
import "core-js/es6/array";
// Poolyfills for old browsers IE9, IE10 e IE11
import * as React from "react";
import "./App.css";
import { CommandBar } from "office-ui-fabric-react/lib/CommandBar";
import IListViewData from "./IListViewData";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import Services from "./services";
import { Pagination } from "antd";
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn
} from "office-ui-fabric-react/lib/DetailsList";
import { Panel, PanelType } from "office-ui-fabric-react/lib/Panel";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Link } from "office-ui-fabric-react/lib/Link";
import { Label } from "office-ui-fabric-react/lib/Label";
import { ActionButton } from "office-ui-fabric-react/lib/Button";
import { SearchBox } from "office-ui-fabric-react/lib/SearchBox";
import IStrings from "./IStrings";
import IAppProps from './IAppProps';
import IAppState from './IAppState';

// Inicializae Office-ui-Fabric Icons
initializeIcons();
var _listViewColumns: IColumn[];
export default class IApp extends React.Component<IAppProps, IAppState> {
  // Private Members
  private _selection: Selection;
  private webUrl = _spPageContextInfo.webAbsoluteUrl;
  private SPDataService = new Services(this.webUrl);
  private strings: IStrings;
  // private _listId: string = this.util.getQueryStringParameter("listId");
  private _listId: string = this.props.ListId;
  // Props
  constructor(props: IAppProps) {
    super(props);
    // State Inicialize
    this.state = {
      listData: [],
      selectedItem: undefined,
      selectedFieldItem: undefined,
      selectedItems: [],
      listViewColumns: [],
      listViewItems: [],
      disableView: true,
      totalListItems: 0,
      currentListPage: 0,
      totalListPages: 0,
      lastPageLoaded: 1,
      showPanel: false,
      selectListItem: {}
    };
    // Set Strings
    this.strings = this.props.Strings;
    // Create ListView Columns
    _listViewColumns = [
      {
        key: "NrMecanografico",
        name: this.strings.nmecanograficoLabel,
        fieldName: "nmecanografico",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
        onColumnClick: this._onColumnClick,
        data: "string",
        onRender: (item: IListViewData) => {
          return <span>{item.nmecanografico}</span>;
        },
        isPadded: true
      },
      {
        key: "NomeColaborador",
        name: this.strings.TitleLabel,
        fieldName: "Title",
        minWidth: 110,
        maxWidth: 125,
        isResizable: true,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick,
        data: "string",
        onRender: (item: IListViewData) => {
          return <span>{item.Title}</span>;
        },
        isPadded: true
      },
      {
        key: "NIFColaborador",
        name: this.strings.nifcolaboradorLabel,
        fieldName: "NIFColaborador",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsable: true,
        isSortedDescending: false,
        data: "string",
        onColumnClick: this._onColumnClick,
        onRender: (item: IListViewData) => {
          return <span>{item.nifcolaborador}</span>;
        },
        isPadded: true
      },
      {
        key: "NomeAluno",
        name: this.strings.nomealunoLabel,
        fieldName: "NomeAluno",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsable: true,
        isSortedDescending: false,
        data: "string",
        onColumnClick: this._onColumnClick,
        onRender: (item: IListViewData) => {
          return <span>{item.nomealuno}</span>;
        }
      },
      {
        key: "DataNascAluno",
        name: this.strings.DataNascAlunoLabel,
        fieldName: "DataNascAluno",
        minWidth: 120,
        maxWidth: 120,
        isResizable: true,
        isCollapsable: true,
        isSortedDescending: false,
        data: "number",
        onColumnClick: this._onColumnClick,
        onRender: (item: IListViewData) => {
          return <span>{item.DataNascAluno}</span>;
        }
      },
      {
        key: "NIFAluno",
        name: this.strings.NIFAlunoLabel,
        fieldName: "NIFAluno",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsable: true,
        isSortedDescending: false,
        data: "number",
        onColumnClick: this._onColumnClick,
        onRender: (item: IListViewData) => {
          return <span>{item.NIFAluno}</span>;
        }
      },
      {
        key: "Holding",
        name: this.strings.HoldingLabel,
        fieldName: "Holding",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsable: true,
        isSortedDescending: false,
        data: "string",
        onColumnClick: this._onColumnClick,
        onRender: (item: IListViewData) => {
          return <span>{item.Holding}</span>;
        }
      }
    ];
    // Get Item Selection
    this._selection = new Selection({
      onSelectionChanged: () => this._getSelectionDetails()
    });
    // Bind Functions
    this._getSelectionDetails = this._getSelectionDetails.bind(this);
    this.OnPageChange = this.OnPageChange.bind(this);
    this.dismissPanel = this.dismissPanel.bind(this);
    this.onView = this.onView.bind(this);
    this.onRefresh = this.onRefresh.bind(this);
    this._onSearchList = this._onSearchList.bind(this);
    this._onSearchBlur = this._onSearchBlur.bind(this);
  }

  // Component Did Mount function
  public async componentDidMount() {
    // Load List Data
    this.loadListData();
  }
  // Render Compoment
  public render() {
    return (
      <div>
        <h1> {this.strings.formatString(this.strings.PageTitle,this.props.Title) }</h1>
        <div style={{ display: "none" }}>
          <p>Total View items {this.state.listViewItems.length}</p>
          <p>Total list Items {this.state.totalListItems} </p>
          <p>Total list Pages {this.state.totalListPages} </p>
        </div>
        <div style={{ display: "inline-block" }}>
          <SearchBox
            placeholder={this.strings.PesquisarPlaceHolder}
            className="ms-SearchBox-field-input"
            onSearch={this._onSearchList}
            onClear={this._onSearchBlur}
            value=""
          />
        </div>
        <div className="CommandBarcontainer">
          <CommandBar
            items={[
              {
                key: "cmdview",
                name: this.strings.CommandViewAction,
                disabled: this.state.disableView,
                iconProps: {
                  iconName: "RedEye"
                },
                onClick: this.onView
              }
            ]}
            farItems={[
              {
                key: "cmdRefresh",
                name: this.strings.commandRefreshAction,
                iconProps: {
                  iconName: "Refresh"
                },
                onClick: this.onRefresh
              }
            ]}
          />
        </div>
        {
          <DetailsList
            items={this.state.listViewItems.slice(
              this.state.currentListPage * this.SPDataService.listPageSize -
                this.SPDataService.listPageSize,
              this.state.currentListPage * this.SPDataService.listPageSize
            )}
            compact={true}
            columns={this.state.listViewColumns}
            selectionMode={SelectionMode.single}
            setKey="set"
            isHeaderVisible={true}
            selection={this._selection}
            enterModalSelectionOnTouch={true}
            layoutMode={DetailsListLayoutMode.justified}
          />
        }
        {// Total Pages > 1 Show Pagination
        this.state.totalListPages > 1 ? (
          <div style={{ marginTop: 20, float: "right" }}>
            <Pagination
              size="small"
              total={this.state.totalListItems}
              pageSize={this.SPDataService.listPageSize}
              current={this.state.currentListPage}
              onChange={this.OnPageChange}
            />
          </div>
        ) : (
          ""
        )}
        <div>
          <Panel
            isOpen={this.state.showPanel}
            onDismiss={this.dismissPanel}
            type={PanelType.medium}
            headerText={this.strings.PanelViewTitle}
          >
            <Label>{this.strings.AnexosLabel}</Label>
            <div
              style={{
                display: "inline-block}",
                marginLeft: 3,
                marginBottom: 15
              }}
            >
              {this.state.selectListItem.attachements != undefined
                ? this.state.selectListItem.attachements.map((file: any) => {
                    return (
                      <ActionButton
                        style={{
                          display: "Inline-block",
                          marginRight: 2
                        }}
                        data-automation-id="Attach"
                        iconProps={{ iconName: "Attach" }}
                        allowDisabledFocus={true}
                        disabled={false}
                        checked={true}
                      >
                        <Link href={file.ServerRelativeUrl}>
                          {file.FileName}
                        </Link>
                      </ActionButton>
                    );
                  })
                : ""}
            </div>
            <div style={{ overflow: "Auto" }}>
              <TextField
                style={{ color: "#454545" }}
                label={this.strings.nmecanograficoLabel}
                disabled={true}
                value={this.state.selectListItem.nmecanografico}
              />
              <TextField
                style={{ color: "#454545" }}
                label={this.strings.TitleLabel}
                disabled={true}
                value={this.state.selectListItem.Title}
              />
              <TextField
                style={{ color: "#454545" }}
                label={this.strings.nifcolaboradorLabel}
                disabled={true}
                value={this.state.selectListItem.nifcolaborador}
              />
              <TextField
                style={{ color: "#454545" }}
                label={this.strings.nomealunoLabel}
                disabled={true}
                value={this.state.selectListItem.nomealuno}
              />
              <TextField
                style={{ color: "#454545" }}
                label={this.strings.NIFAlunoLabel}
                disabled={true}
                value={this.state.selectListItem.NIFAluno}
              />
              <TextField
                style={{ color: "#454545" }}
                label={this.strings.DataNascAlunoLabel}
                disabled={true}
                value={this.state.selectListItem.DataNascAluno}
              />
              <TextField
                style={{ color: "#454545" }}
                label={this.strings.IdadeAlunoLabel}
                disabled={true}
                value={this.state.selectListItem.IdadeAluno}
              />
              <TextField
                style={{ color: "#454545" }}
                label={this.strings.HoldingLabel}
                disabled={true}
                value={this.state.selectListItem.Holding}
              />
              <TextField
                style={{ color: "#454545" }}
                label={this.strings.areaLabel}
                disabled={true}
                value={this.state.selectListItem.area}
              />
              <TextField
                style={{ color: "#454545" }}
                label={this.strings.codigopostalLabel}
                disabled={true}
                value={this.state.selectListItem.codigopostal}
              />
              <TextField
                style={{ color: "#454545" }}
                label={this.strings.localidadeLabel}
                disabled={true}
                value={this.state.selectListItem.localidade}
              />
              <TextField
                style={{ color: "#454545" }}
                label={this.strings.anoLabel}
                disabled={true}
                value={this.state.selectListItem.ano}
              />
              <TextField
                style={{ color: "#454545" }}
                label={this.strings.mediaLabel}
                disabled={true}
                value={this.state.selectListItem.media}
              />
              <TextField
                style={{ color: "#454545" }}
                label={this.strings.moradaLabel}
                disabled={true}
                value={this.state.selectListItem.morada}
              />
              <TextField
                style={{ color: "#454545" }}
                label={this.strings.empresaLabel}
                disabled={true}
                value={this.state.selectListItem.empresa}
              />
              <TextField
                style={{ color: "#454545" }}
                label={this.strings.nomelojaLabel}
                disabled={true}
                value={this.state.selectListItem.nomeloja}
              />
              <TextField
                style={{ color: "#454545" }}
                label={this.strings.consentimentoLabel}
                disabled={true}
                value={this.state.selectListItem.consentimento}
              />
            </div>
          </Panel>
        </div>
      </div>
    );
  }
  // Show Detail Panel
  private onRefresh(ev: React.MouseEvent<HTMLElement>) {
    ev.preventDefault();
    // Clear any Selection first
    this._selection.setItems(this.state.listViewItems, true);
    this.loadListData();
  }
  // Show Detail Panel
  private onView(ev: React.MouseEvent<HTMLElement>) {
    ev.preventDefault();
    this.setState({ showPanel: true });
  }
  // Dissmiss Panel
  private dismissPanel() {
    this.setState({ showPanel: false });
  }
  // Page Navigation Changed
  public async OnPageChange(page: number) {
    let { lastPageLoaded, listViewItems, currentListPage } = this.state;
    // Testa se pagina de items já carregada
    currentListPage = page;
    if (currentListPage > lastPageLoaded) {
      // pagina de items ainda não carregada do SharePoint
      let _numeroPagesToload = page - lastPageLoaded;
      for (let index = 1; index <= _numeroPagesToload; index++) {
        // If page data not loaded yet load it
        const SPListData: IListViewData[] = await this.SPDataService.GetPageItems();
        listViewItems = [...listViewItems, ...SPListData];
      }
      lastPageLoaded = page;
    }
    // Upate State
    this.setState({
      lastPageLoaded: lastPageLoaded,
      listViewItems: listViewItems,
      currentListPage: currentListPage
    });
  }
  //    Obter Item Seleccionado da ListView e activa/desactiva opções
  private async _getSelectionDetails() {
    const selectionCount = this._selection.getSelectedCount();
    if (selectionCount !== 0) {
      let selectItem = this._selection.getSelection()[0] as IListViewData;
      let _itemId = parseInt(selectItem.key);
      let _attachements: any[] = await this.SPDataService.GetListItemAttachments(
        this._listId,
        _itemId
      );
      selectItem.attachements = _attachements;
      this.setState({ disableView: false, selectListItem: selectItem });
    } else {
      this.setState({ disableView: true });
    }
  }
  // Search Items
  private async _onSearchList(value: string) {
    if (value.trim() == "") return;
    let { listViewItems, currentListPage, totalListItems } = this.state;
    const SPListData: IListViewData[] = await this.SPDataService.searchList(
      this._listId,
      value
    );
    totalListItems = SPListData.length;
    listViewItems = SPListData;
    currentListPage = 1;
    let _listTotalPages: number = 0;
    _listTotalPages = await this.getTotalListPages(totalListItems);
    // Clear any Selection first
    this._selection.setItems(listViewItems, true);
    // Update State
    this.setState({
      totalListItems: totalListItems,
      listViewItems: listViewItems,
      currentListPage: currentListPage,
      totalListPages: _listTotalPages
    });
  }

  // On Blur SearchBox
  private _onSearchBlur(ev: any) {
    ev.preventDefault();
    // Refresh List
    this.loadListData();
  }
  // Load List Data Initial load or Refresh
  private async loadListData() {

    const SPListData: IListViewData[] = await this.SPDataService.GetListData(
      this._listId
    );
    const _totalListItems: number = await this.SPDataService.GetTotalItems(
      this._listId
    );
    let _listTotalPages: number = 0;
    _listTotalPages = await this.getTotalListPages(_totalListItems);
    // "4cf37415-0050-497d-b22b-ba8273a751b9"
    // const SPListData:IListData[] = await SPDataService.GetListData("C7937559%2D924A%2D48B5%2D9363%2DA30352428E87")

    this.setState({
      listViewItems: SPListData,
      listViewColumns: _listViewColumns,
      totalListPages: _listTotalPages,
      totalListItems: _totalListItems,
      currentListPage: 1
    });
  }

  // Calcula Total List Pages
  private async getTotalListPages(_totalListItems: number): Promise<number> {
    let _listTotalPages: number = 0;
    // Get Total Pages of list
    if (_totalListItems > 0) {
      _listTotalPages = Math.floor(
        _totalListItems / this.SPDataService.listPageSize
      );
      _listTotalPages =
        _listTotalPages == 0 ? _listTotalPages + 1 : _listTotalPages;
      // Get Remainder
      if (_totalListItems > this.SPDataService.listPageSize) {
        let _remainderValue = Math.floor(
          _totalListItems % this.SPDataService.listPageSize
        );
        _listTotalPages =
          _remainderValue == 0 ? _listTotalPages : _listTotalPages + 1;
      }
    }
    return _listTotalPages;
  }
  // Sort Columns:
  private _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const { listViewColumns, listViewItems } = this.state;
    let newItems: IListViewData[] = listViewItems.slice();
    const newColumns: IColumn[] = listViewColumns.slice();
    const currColumn: IColumn = newColumns.filter(
      (currCol: IColumn, idx: number) => {
        return column.key === currCol.key;
      }
    )[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    newItems = this._sortItems(
      newItems,
      currColumn.fieldName,
      currColumn.isSortedDescending
    );
    this.setState({
      listViewColumns: newColumns,
      listViewItems: newItems
    });
  };
  // Sort List Items
  private _sortItems = (
    items: IListViewData[],
    sortBy: string,
    descending = false
  ): IListViewData[] => {
    if (descending) {
      return items.sort((a: IListViewData, b: IListViewData) => {
        if (a[sortBy] < b[sortBy]) {
          return 1;
        }
        if (a[sortBy] > b[sortBy]) {
          return -1;
        }
        return 0;
      });
    } else {
      return items.sort((a: IListViewData, b: IListViewData) => {
        if (a[sortBy] < b[sortBy]) {
          return -1;
        }
        if (a[sortBy] > b[sortBy]) {
          return 1;
        }
        return 0;
      });
    }
  };
}
