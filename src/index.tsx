// Poolyfills for old browsers IE10 and IE11
import 'core-js/es6/map';
import 'core-js/es6/set';
import 'raf/polyfill';
import 'core-js/es6/promise'
import 'core-js/es6/array'
// Poolyfills for old browsers IE10 and IE11
import * as React from "react";
import * as ReactDOM from "react-dom";
import App from "./App";
import "./index.css";
import registerServiceWorker from "./registerServiceWorker";
import Util from "./Utils";
import Services from "./services";
import IStrings from './IStrings';

let util = new Util();
let strings: IStrings;
let webUrl = _spPageContextInfo.webAbsoluteUrl;
let SPDataService = new Services(webUrl);
// Get Parameter
let _listId: string = util.getQueryStringParameter("listId");
// Get LCID Strings
util.getLCID().then( stringsLCID => {
  strings = stringsLCID;
});
// Get List Title
SPDataService.GetListITitle(_listId)
  .then(value => {
    ReactDOM.render(
      <App ListId={_listId} Title={value} Strings={strings}/>,
      document.getElementById("root") as HTMLElement
    );
  })
  .catch((reason: any) => {
    ReactDOM.render(
      <div>
        <p>
           {reason}
        </p>
      </div>,
      document.getElementById("root") as HTMLElement
    );
  });
  registerServiceWorker();
