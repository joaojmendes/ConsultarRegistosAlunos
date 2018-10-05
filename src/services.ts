/**
 * Comment:
 *
 * Date:   14/9/2018
 * Author: João Mendes
 * Company: Site Na Nuvem
 *
 */
import IListData from "./IListData";
import IListViewData from "./IListViewData";
import * as jquery from "jquery";
import { sp } from "@pnp/sp";
import * as moment from "moment";
// Class Data Services
export default class Services {
  // Private Class Vars
  private wrkListData: any = [];
  private _listPageSize: number = 10;
  private webUrl: string = "";
  public listPageSize: number = this._listPageSize;
  // Constructor
  constructor(siteUrl: string) {
    this.webUrl = siteUrl;
    // Configure PnPjs
    sp.setup({
      sp: {
        baseUrl: this.webUrl,
        headers: {
          Accept: "application/json;odata=verbose"
        }
      }
    });
  }
  /**
   * Get List Data
   * listID = GUID da Lista a ler
   */
  public async GetListData(listId: string): Promise<IListViewData[]> {
    // const web = new sp.web(webUrl);
    let _userHolding: any = await this.GetHoldingForUser();
    this.wrkListData = await sp.web.lists
      .getById(listId)
      .items.filter("Holding eq " + _userHolding.Id)
      .orderBy("Id")
      .top(this._listPageSize)
      .getPaged();
    let listData: IListViewData[] = [];
    this.wrkListData.results.map((element: IListData) => {
      listData.push({
        key: element.Id.toString(),
        Title: element.Title.toString(),
        NIFAluno: element.NIFAluno == undefined ? "" : element.NIFAluno,
        nifcolaborador:
          element.nifcolaborador == undefined
            ? ""
            : element.nifcolaborador.toString(),
        nomealuno:
          element.nomealuno == undefined ? "" : element.nomealuno.toString(),
        DataNascAluno:
          element.DataNascAluno == undefined
            ? ""
            : moment(element.DataNascAluno.toString())
                .format("dddd, MMMM Do YYYY")
                .toString(),
        IdadeAluno:
          element.IdadeAluno == undefined ? "" : element.IdadeAluno.toString(),
        Holding: _userHolding.Title,
        nmecanografico:
          element.nmecanografico == undefined
            ? ""
            : element.nmecanografico.toString(),
        area: element.area == undefined ? "" : element.area.toString(),
        codigopostal:
          element.codigopostal == undefined
            ? ""
            : element.codigopostal.toString(),
        localidade:
          element.localidade == undefined ? "" : element.localidade.toString(),
        media: element.media == undefined ? "" : element.media.toString(),
        morada: element.morada == undefined ? "" : element.morada.toString(),
        empresa: element.empresa == undefined ? "" : element.empresa.toString(),
        nomeloja:
          element.nomeloja == undefined ? "" : element.nomeloja.toString(),
        consentimento:
          element.consentimento == undefined
            ? ""
            : element.consentimento.results[0],
        ano: element.ano.toString() == undefined ? "" : element.ano.toString()
      });
    });
    return listData;
  }
  /*
  // Get NextPage List Items
  */
  public async GetPageItems(): Promise<IListViewData[]> {
    let listData: IListViewData[] = [];
    if (this.wrkListData.hasNext) {
      let _userHolding: any = await this.GetHoldingForUser();
      // this will carry over the type specified in the original query for the results array
      this.wrkListData = await this.wrkListData.getNext();
      this.wrkListData.results.map((element: IListData) => {
        listData.push({
          key: element.Id.toString(),
          Title: element.Title.toString(),
          NIFAluno: element.NIFAluno == undefined ? "" : element.NIFAluno,
          nifcolaborador:
            element.nifcolaborador == undefined
              ? ""
              : element.nifcolaborador.toString(),
          nomealuno:
            element.nomealuno == undefined ? "" : element.nomealuno.toString(),
          DataNascAluno:
            element.DataNascAluno == undefined
              ? ""
              : moment(element.DataNascAluno.toString())
                  .format("dddd, MMMM Do YYYY")
                  .toString(),
          IdadeAluno:
            element.IdadeAluno == undefined
              ? ""
              : element.IdadeAluno.toString(),
          Holding: _userHolding.Title,
          nmecanografico:
            element.nmecanografico == undefined
              ? ""
              : element.nmecanografico.toString(),
          area: element.area == undefined ? "" : element.area.toString(),
          codigopostal:
            element.codigopostal == undefined
              ? ""
              : element.codigopostal.toString(),
          localidade:
            element.localidade == undefined
              ? ""
              : element.localidade.toString(),
          media: element.media == undefined ? "" : element.media.toString(),
          morada: element.morada == undefined ? "" : element.morada.toString(),
          empresa:
            element.empresa == undefined ? "" : element.empresa.toString(),
          nomeloja:
            element.nomeloja == undefined ? "" : element.nomeloja.toString(),
          consentimento:
            element.consentimento == undefined
              ? ""
              : element.consentimento.toString(),
          ano: element.ano == undefined ? "" : element.ano.toString()
        });
      });
    }
    return listData;
  }
  // Get Holding for current user
  private async GetHoldingForUser(): Promise<string> {
    const listId: string = "8ad001b5-061d-42c4-8ade-67bf7c99586d"; // mudar para o ID da Lista de Holdings
    let _listItem = await sp.web.lists
      .getById(listId) // ListId de Pivots e Holding
      .items.select("Pivot/Id,Title,Id")
      .expand("Pivot")
      .filter("Pivot/Id eq " + _spPageContextInfo.userId)
      .get();
    return _listItem[0];
  }
  // Get Total Of List Items
  public async GetTotalItems(listId: string) {
    let url: string =
      this.webUrl + "/_api/web/lists('" + listId + "')/ItemCount";
    return new Promise<number>((resolve, rejectd) => {
      jquery.ajax({
        url: url,
        dataType: "json",
        headers: {
          accept: "application/json; odata=verbose",
          "content-type": "application/json;odata=verbose"
        },
        success: function(data: any) {
          resolve(data.d.ItemCount);
        },
        error: function(xhr: any, errorType: any, exception: any) {
          rejectd(0);
        }
      });
    });
  }

  /*
   Get List Item Attachments
  */
  public async GetListItemAttachments(
    listId: string,
    itemId: number
  ): Promise<any[]> {
    let _attachments = await sp.web.lists
      .getById(listId)
      .items.getById(itemId)
      .attachmentFiles.get();
    return _attachments;
  }
  /*
  // Get List Item Title
  */
  public async GetListITitle(listId: string): Promise<string> {
    let _listTitle = await sp.web.lists.getById(listId).get();
    return _listTitle.Title;
  }
  /*
  **Search List
  */
  public async searchList(
    listId: string,
    value: string
  ): Promise<IListViewData[]> {
    let _userHolding: any = await this.GetHoldingForUser();
    let _searchString = String.format(
      "(startswith(nomealuno,'{0}')) or (startswith(nmecanografico,'{0}')) or (startswith(Title,'{0}'))  or (startswith(NIFAluno,'{0}')) or (startswith(nifcolaborador,'{0}'))",
      value
    );
    this.wrkListData = await sp.web.lists
      .getById(listId)
      .items.filter(_searchString)
      .orderBy("Id")
      .top(5000)
      .getPaged();
    let listData: IListViewData[] = [];
    this.wrkListData.results.map((element: IListData) => {
      listData.push({
        key: element.Id.toString(),
        Title: element.Title.toString(),
        NIFAluno: element.NIFAluno == undefined ? "" : element.NIFAluno,
        nifcolaborador:
          element.nifcolaborador == undefined
            ? ""
            : element.nifcolaborador.toString(),
        nomealuno:
          element.nomealuno == undefined ? "" : element.nomealuno.toString(),
        DataNascAluno:
          element.DataNascAluno == undefined
            ? ""
            : moment(element.DataNascAluno.toString())
                .format("dddd, MMMM Do YYYY")
                .toString(),
        IdadeAluno:
          element.IdadeAluno == undefined ? "" : element.IdadeAluno.toString(),
        Holding: _userHolding.Title,
        nmecanografico:
          element.nmecanografico == undefined
            ? ""
            : element.nmecanografico.toString(),
        area: element.area == undefined ? "" : element.area.toString(),
        codigopostal:
          element.codigopostal == undefined
            ? ""
            : element.codigopostal.toString(),
        localidade:
          element.localidade == undefined ? "" : element.localidade.toString(),
        media: element.media == undefined ? "" : element.media.toString(),
        morada: element.morada == undefined ? "" : element.morada.toString(),
        empresa: element.empresa == undefined ? "" : element.empresa.toString(),
        nomeloja:
          element.nomeloja == undefined ? "" : element.nomeloja.toString(),
        consentimento:
          element.consentimento == undefined
            ? ""
            : element.consentimento.results[0],
        ano: element.ano.toString() == undefined ? "" : element.ano.toString()
      });
    });
    return listData;
  }
}
