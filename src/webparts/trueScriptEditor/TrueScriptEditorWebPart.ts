import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneTextFieldProps
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './TrueScriptEditor.module.scss';
import * as strings from 'trueScriptEditorStrings';
import { ITrueScriptEditorWebPartProps } from './ITrueScriptEditorWebPartProps';
import * as jQuery from 'jQuery';
export default class TrueScriptEditorWebPart extends BaseClientSideWebPart<ITrueScriptEditorWebPartProps> {

  public render(): void {
    window['trueScriptEditor'] = this;
    window['jQuery'] = jQuery;
    var loaded = [];
    jQuery('head').append('<style truescript></style>');
    jQuery('link[truescript]').each(function(){ jQuery(this).remove(); });
    jQuery(this.domElement).empty();
    if (this.properties.cssEditor || this.properties.extCSS || this.properties.extJS || this.properties.scriptEditor || this.properties.htmlEditor) {
      if (this.properties.scriptEditor){
        window['jQuery'] = jQuery;
        eval(this.properties.scriptEditor);
      }
      if (this.properties.extJS){
        this.properties.extJS.split(',').forEach(function(scr){
          if (loaded.indexOf(scr) <= -1){
            loaded.push(scr);
            window['jQuery'].getScript(scr);
          }
        });
      }
      if (this.properties.cssEditor){
        jQuery('style[truescript]').empty().html(this.properties.cssEditor);
      }
      if (this.properties.extCSS){
        jQuery('link[truescript]').each(function(){
          jQuery(this).remove();
        });
        this.properties.extCSS.split(',').forEach(function(css){
          jQuery('head').append('<link truescript rel="stylesheet" href="' + css + '" type="text/css" />');
        });
      }
      if (this.properties.htmlEditor){
        jQuery(this.domElement).html(this.properties.htmlEditor);
      }      
    } else {
    this.domElement.innerHTML = `
      <div class="${styles.helloWorld}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">True Script Editor Web Part</span>
              <p class="ms-font-l ms-fontColor-white">Edit the web part properties to add your own CSS or Javascript.</p>
              <p class="ms-font-l ms-fontColor-white">jQuery is available using the 'jQuery' prefix ($ not supported).</p>
              <a href="javascript:window['trueScriptEditor']._context._propertyPaneAccessor.open()" class="${styles.button}">
                <span class="${styles.label}">Open Properties</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
    }
    
  }

  public editProperties(){
    this.context.propertyPane.open();
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('extJS',{
                  label: 'External JS Files (Comma Seperated)',
                  multiline: true,
                  placeholder: 'Use this input to load external JavaScript Files, type or paste values in a commas seperated list *without* spaces'
                }),
                PropertyPaneTextField('scriptEditor',{
                  label: 'Script Editor',
                  multiline: true,
                  placeholder: 'Javascript in this field will be executed, **Do not include <script> tags.'
                }),
                PropertyPaneTextField('extCSS',{
                  label: 'External CSS Files (Comma Seperated)',
                  multiline: true,
                  placeholder: 'Use this input to load external CSS Files, type or paste values in a commas seperated list *without* spaces'
                }),
                PropertyPaneTextField('cssEditor',{
                  label: 'CSS Editor',
                  multiline: true,
                  placeholder: 'CSS in this field will be added to the page, **Do not include <style> tags.'
                }),
                PropertyPaneTextField('htmlEditor',{
                  label: 'HTML Editor',
                  multiline: true,
                  placeholder: 'HTML in this field will be added to the page, **Do not include <html>,<body> nor <footer> tags.'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}