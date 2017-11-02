import { Log } from '@microsoft/sp-core-library';
import { LogHandler, LogLevel } from '../../common/LogHandler';

import { override } from '@microsoft/decorators';
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';

import * as strings from 'UrlLinkFieldCustomizerStrings';
import styles from './UrlLinkFieldCustomizer.module.scss';

import pnp from "sp-pnp-js";

/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IUrlLinkFieldCustomizerProperties {
  // A target attribute specification
  target?: string;
}

const LOG_SOURCE: string = 'UrlLinkFieldCustomizer';

export default class UrlLinkFieldCustomizer
  extends BaseFieldCustomizer<IUrlLinkFieldCustomizerProperties> {

  //Instead of querying the column description for EVERY row, we use this object to hold a single promise for eahc 
  pnpAllPromises = new Map<string, Promise<any>>();

  @override
  public onInit(): Promise<void> {
    // Add your custom initialization to this method.  The framework will wait
    // for the returned promise to resolve before firing any BaseFieldCustomizer events.
    Log._initialize(new LogHandler((window as any).LOG_LEVEL || LogLevel.Verbose));
    Log.info(LOG_SOURCE, 'Activated UrlLinkFieldCustomizer with properties:');
    //Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    Log.info(LOG_SOURCE, `The following string should be equal: "UrlLinkFieldCustomizer" and "${strings.Title}"`);
    
    //return Promise.resolve();
    //-
    //Get pnp Context -- per documentation this was required
    var promise = super.onInit().then(_ => {
      pnp.setup({
        spfxContext: this.context
      });
    });
    return promise;
  }

  @override
  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    /*
    // Use this method to perform your custom cell rendering.
    const text: string = `${this.properties.sampleText}: ${event.fieldValue}`;
    event.domElement.innerText = text;
    event.domElement.classList.add(styles.cell);
    */

    // This method performs our custom cell rendering.
    // Called to render each cell [in the list] that we are customizing
    // -Step 1: Using pnp, lookup the column description for what we are rendering -- this is expected to be a URL
    // -Step 2: [not using pnp] Get the field value for the current row -- this is expected to be a JSON string
    // -Step 3: Iterate through all the JSON properties and replace appropriate tokens in URL. 
    //          Thus, we paramertize the URL and replace it with values from the JSON for the individual row.
    // -Step 4: Create <a href>Title</a> syntax and render for cell. NOTE: The "Title" field in the JSON defines text between <a> and </a>. 
    //          If not present, the URL is not rendered.
    //
    //NOTE: I'd rather not use pnp, but it appears to be the only way to retrieve the column description. That
    //      value is not available through the event nor this objects.
    
    //Get ListId/FieldId of current List/Field we are to render. We will use this with 
    //pnp to lookup the field's description -- which is expected to be a URL.
    var listID = this.context.pageContext.list.id.toString();
    var fieldID = this.context.field.id.toString();
    let fieldDescription: string;

    let pnpPromise:Promise<any>;

    if(this.pnpAllPromises.has(fieldID))
    {
      //We already created a promise, get that one and don't create a new one
      pnpPromise = this.pnpAllPromises.get(fieldID);
      //Log.info(LOG_SOURCE, "Reduced work. Retrieved existing promise " + fieldID + ".");
    }
    else
    {
      //Create new promise and add to list
      pnpPromise = pnp.sp.web.lists.getById(listID).fields.getById(fieldID).get();
      this.pnpAllPromises.set(fieldID, pnpPromise);
      //Log.info(LOG_SOURCE, "Created NEW promise " + fieldID + ".");
    }

    //Use pnp to get information on the column we are rendering. Again, we are looking for the column description.
    //pnp.sp.web.lists.getById(listID).fields.getById(fieldID).get().then((item: any) => {
    pnpPromise.then((item: any) => {

      //Expecting fieldDescription to be a URL
      fieldDescription = item.Description.toString();
      //Log.info(LOG_SOURCE, "Field Description: " + fieldDescription);

      //Get cell's [current row/column] value. We expect it to be a JSON string.
      //In SharePoint, one might expect the Calculated Column to be defined as:
      //   ="{  ""Title"" : ""ClickMe"", " & "  ""DocID"" : """ & [Document ID Value] & """}"
      let columnData: any;
      try {
        //Log.info(LOG_SOURCE, "FieldValue: " + event.fieldValue);
        columnData = JSON.parse(event.fieldValue);
      }
      catch (e) {
        Log.warn(LOG_SOURCE, "Exception parsing JSON for column. " + e);
      }

      //Iterate through the JSON object hopefully we just got and then do a find/replace on any parameters in the URL string
      try
      {
        Object.keys(columnData).forEach(function (key, index) {
          // key: the name of the object key
          // index: the ordinal position of the key within the object 

          var regex = new RegExp("[{]" + key + "[}]", "g");
          fieldDescription = fieldDescription.replace(regex, columnData[key]);
        });
        //Log.info(LOG_SOURCE, "Updated URL: " + fieldDescription);
      }
      catch(e) {
        Log.warn(LOG_SOURCE, "Exception replacing URL elements from JSON object. " + e);
      }

      //Build-up HTML with manipulated URL from field description
      try {
        
        //Only show URL link if Title is specified
        if(columnData.Title != undefined)
        {
          //I built up the using dom elements to better support complicated URLs that fail when trying to 
          // assemble them with string concatenations.
          var divElement = document.createElement("div");
          //divElement.setAttribute("class", styles.cell);
          var aElement = document.createElement("a");
          //Log.info(LOG_SOURCE, "Properties: " + this.properties);
          if(this.properties.target != undefined)
            aElement.target = this.properties.target; // = "_blank";
          else
            aElement.target = "_blank";
          //-
          aElement.href = fieldDescription;
        
          var aText = document.createTextNode(columnData.Title);
          aElement.appendChild(aText);
          divElement.appendChild(aElement);
          //-
          event.domElement.innerHTML = divElement.outerHTML;
        }
        else
        {
          Log.warn(LOG_SOURCE, "No Title specified in JSON (" + event.fieldValue + ") for column row.");
        }
      }
      catch (e) {
        Log.warn(LOG_SOURCE, "Exception forming clickable URL. " + e);
      }
    });
  }

  @override
  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    super.onDisposeCell(event);
  }
}
