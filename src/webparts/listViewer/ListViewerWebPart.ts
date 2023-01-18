import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ListViewerWebPart.module.scss';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { IColumnReturnProperty, IPropertyFieldRenderOption, PropertyFieldColumnPicker, PropertyFieldColumnPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldColumnPicker';
import * as strings from 'ListViewerWebPartStrings';
import { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';
import { PropertyPaneMarkdownContent } from '@pnp/spfx-property-controls/lib/PropertyPaneMarkdownContent';
import { PropertyFieldMessage} from '@pnp/spfx-property-controls/lib/PropertyFieldMessage';
import { MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';


import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/files";

import "@pnp/sp/profiles";  

export interface IListViewerWebPartProps {
  //default in render
  script: string;
  //no default
  list: string;
  //default in manifest
  multiColumn: string[];
  outerHtml: string;
  itemHtml: string;
}

export default class ListViewerWebPart extends BaseClientSideWebPart<IListViewerWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    if(this.properties.script == undefined){this.properties.script = `
    
    //window["test"]=t
t.domElement.innerHTML = t.properties.outerHtml;

let container = t.domElement.querySelector("[replaceItems=true]")

if(container!==null)
//get Department
t.getCurrentUserProperty("Department").then((f) => {
    
//get List Items
t.getListItems(
        t.properties.list,
        t.properties.multiColumn,
        12,"").then((e)=> {
            //console.log(e)
for(let i=0;i<e.length;i++){
    let item = e[i];
    let element = t.buildHTML(t.properties.itemHtml,item);
    container.appendChild(element);
}
    
});
    
});
    
    
    
    `
    }
    this.runFunction();
    
  }
  async getListItems(listId:string,columns:string[],numItems:number,query:string="",orderField:string="Title"){
    var _sp = spfi().using(SPFx(this.context));
    const r = await _sp.web.lists.getById(listId).items //this.properties.list
        .select(columns.toString())//this.properties.fields
        .filter(query)//this.properties.query "Title eq '" + "dpt Test" + "'"
        //.expand(this.properties.expand)
        .top(numItems)//parseInt(this.properties.top)
        .orderBy(orderField,true)();
    return r;
  }
  async runFunction(){
    if(this.properties.script==undefined)this.properties.script="";
    const AsyncFunction = Object.getPrototypeOf(async function(){}).constructor;
    let f = new AsyncFunction("t",this.properties.script);
    f(this);
  }
  async getCurrentUserProperty(prop:string){//TODO
    let a = await this.getCurrentUserProperties([prop]);
    return a[prop];
  }
  async getCurrentUserProperties(props:string[]){
    var _sp = spfi().using(SPFx(this.context));
    const profileS = await _sp.profiles.userProfile;  
    const loginName = profileS.AccountName;  
    const profile = await _sp.profiles.getPropertiesFor(loginName);  
    let profileProp =profile.UserProfileProperties
    let propValue:any={}; let i=0;
    while(Object.keys(propValue).length!==props.length && i<profileProp.length){
      const currentProp = profileProp[i];
      for(let j=0;j<props.length;j++){
        let prop = props[j];
        if(currentProp.Key===prop)
          propValue[currentProp.Key]=currentProp.Value;
      }
      i++;
    }
    return propValue;
  }
  buildHTML (html:string,properties:string[]){
    let customHTML=html;
    for (const key in properties) {
      if (Object.prototype.hasOwnProperty.call(properties, key)) {
        const propertyValue = properties[key];
        const re = new RegExp(`Replace${key}`, 'g');
        customHTML = customHTML.replace(re, propertyValue);
      }
    }
    let returnElement = document.createElement("div")
    returnElement.innerHTML = customHTML
    return returnElement;
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();
    return super.onInit();
  }
  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }
    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }
  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldListPicker('list', {
                  label: 'Select a list',
                  selectedList: this.properties.list,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  // @ts-ignore
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                }),
                // Multi column selection returning the 'Title' of the selected columns as a multi-select dropdown
                PropertyFieldColumnPicker('multiColumn', {
                  label: 'Select columns',
                  // @ts-ignore
                  context: this.context,
                  selectedColumn: this.properties.multiColumn,
                  listId: this.properties.list,
                  disabled: false,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'multiColumnPickerFieldId',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Title"],
                  multiSelect: true,
                  renderFieldAs: IPropertyFieldRenderOption["Multiselect Dropdown"],
                  columnsToExclude: ["Created By","Modified By","Version","Item Child Count","Folder Child Count","Label setting","Retention label","Retention label Applied","Label applied by","App Created By","App Modified By","Item is a Record"]
                }),
                PropertyFieldCodeEditor('outerHtml', {
                  label: 'Outer HTML',
                  panelTitle: 'Edit HTML Code',
                  initialValue: this.properties.outerHtml,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  key: 'codeEditorFieldId',
                  language: PropertyFieldCodeEditorLanguages.HTML,
                  options: {
                    wrap: true,
                    fontSize: 12,
                    // more options
                  }
                }),
                PropertyFieldCodeEditor('itemHtml', {
                  label: 'Item HTML',
                  panelTitle: 'Edit HTML Code',
                  initialValue: this.properties.itemHtml,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  key: 'codeEditorFieldId',
                  language: PropertyFieldCodeEditorLanguages.HTML,
                  options: {
                    wrap: true,
                    fontSize: 12,
                    // more options
                  }
                }),
                PropertyFieldMessage("", {
                  key: "MessageKey",
                  text: `html loads as type string on properties`,
                  messageType: MessageBarType.success,
                  isVisible: true,
                  //multiline: true,
                  //truncate: true
                }),
                PropertyFieldCodeEditor('script', {
                  label: 'onRender',
                  panelTitle: 'Edit JS Code',
                  initialValue: this.properties.script,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  key: 'codeEditorFieldId',
                  language: PropertyFieldCodeEditorLanguages.JavaScript,
                  options: {
                    wrap: true,
                    fontSize: 12,
                    // more options
                  }
                }),
                PropertyPaneMarkdownContent({
                  markdown: `<div style="background-color: #dff6dd;padding:8px;border-radius: 6px;">
                                function to call onRender<br>
                                t is the webpart object
                                <fieldset>
                                  <legend>t.properties:</legend>
                                    list: string;<br>
                                    multiColumn: string[];<br>
                                    outerHtml: string;<br>
                                    itemHtml: string;<br>
                                </fieldset>
                                <fieldset>
                                  <legend>functions:</legend>
                                    async getCurrentUserProperties(
                                      <div style="padding:8px;padding-top: 0px;">
                                        properties:string[])<br>
                                      </div>
                                    async getCurrentUserProperty(
                                      <div style="padding:8px;padding-top: 0px;">
                                        property:string)<br>
                                      </div>
                                    async getListItems(
                                      <div style="padding:8px;padding-top: 0px;">
                                        listId:string,<br>
                                        columns:string[],<br>
                                        numItems:number,<br>
                                        query:string="",<br>
                                        orderField="Title")
                                      </div>
                                    buildHTML(
                                      <div style="padding:8px;padding-top: 0px;">
                                        html:string,<br>
                                        properties:string[])<br>
                                        //returns a domElement
                                      </div>
                                </fieldset>
                              </div>
                  `,
                  key: 'markdownSample',
                  
                  options: {
                    //disableParsingRawHTML: false,
                    namedCodesToUnicode: {
                        le: '\u2264',
                        ge: '\u2265',
                    }
                  }}),
              ]
            }
          ]
        }
      ]
    };
  }
}