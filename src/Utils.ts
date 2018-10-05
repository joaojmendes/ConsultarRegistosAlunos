
// Utils Class
// to com
// to Common Functions

import LocalizedStrings from "react-localization";
import IStrings from "./IStrings";
var locales = require("./locales/locales.json");

export default class Utils {
  private strings: IStrings;
  // Get Query String Parameter
  public getQueryStringParameter(paramToRetrieve: string): string {
    let params = document.URL.split("?")[1].split("&");
    for (var i = 0; i < params.length; i = i + 1) {
      var singleParam = params[i].split("=");
      if (singleParam[0] == paramToRetrieve) return singleParam[1].toString();
    }
    return "";
  }
  public getCurrentUserEmail() {
    return new Promise<string>((Resolve, Reject) => {
      let _email = _spPageContextInfo.userEmail;
      Resolve(_email);
    });
  }
  /*
  Get LCID for the App
  */
  public getLCID() {
    return new Promise<IStrings>((Resolve, Reject) => {
      // Set Strings to locale retrive lcid of SharePoint Site
      this.strings = new LocalizedStrings(locales);
      if (this.strings.getAvailableLanguages().length == 0) {
        Reject(this.strings);
      } else {
        const _hasLCID = this.strings
          .getAvailableLanguages()
          .find(
            (value: string) =>
              value == _spPageContextInfo.currentLanguage.toString()
          );
        /* Teste if current site LCID is define in Locales
    */
        _hasLCID != undefined
          ? this.strings.setLanguage(
              _spPageContextInfo.currentLanguage.toString()
            )
          : this.strings.setLanguage("1033");
        Resolve(this.strings);
      }
    });
  }
}
