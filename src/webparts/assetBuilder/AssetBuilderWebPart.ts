import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneLabel,
} from '@microsoft/sp-webpart-base';

import * as strings from 'AssetBuilderWebPartStrings';
import AssetBuilder from './components/AssetBuilder';
import { IAssetBuilderProps } from './components/IAssetBuilderProps';

//  >>>> ADD import additional controls/components
import { sp, } from "@pnp/sp";
import { UrlFieldFormatType, Field } from "@pnp/sp/presets/all";
import { IFieldAddResult } from "@pnp/sp/fields/types";
import { ChoiceFieldFormatType } from "@pnp/sp/fields/types";
import { DateTimeFieldFormatType, FieldTypes } from "@pnp/sp/fields/types";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";

export interface IAssetBuilderWebPartProps {
  description: string;
  localListName: string;
  localListConfirmed: boolean;
}

const LOG_SOURCE: string = 'RedirectApplicationCustomizer';



export default class AssetBuilderWebPart extends BaseClientSideWebPart<IAssetBuilderWebPartProps> {

  @override
  public async onInit(): Promise<void> {

    // initialize PnP JS library to play with SPFx contenxt
    sp.setup({
      spfxContext: this.context
    });

    // read the server relative URL of the current page from Legacy Page Context
    const currentPageRelativeUrl: string = this.context.pageContext.legacyPageContext.serverRequestPath;
/*
    if (await this.ensureSocialiis7List()) {
      alert('Done Ensuring!');
    }
*/

  }

  // this method ensures that the Redirections lists exists, or if it doesn't exist
  // it creates it, as long as the currently connected user has proper permissions
  private async ensureSocialiis7List(myListName: string, myListDesc: string): Promise<boolean> {
    
    let result: boolean = false;

    try {
      const ensureResult = await sp.web.lists.ensure(myListName,
        myListDesc,
        100,
        true);

      // if we've got the list
      if (ensureResult.list != null) {
        // if the list has just been created
        if (ensureResult.created) {
          // we need to add the custom fields to the list
          //https://pnp.github.io/pnpjs/sp/lists/#ensure-that-a-list-exists-by-title
          //https://pnp.github.io/pnpjs/sp/fields/

          //Add this after creating field to change title:  //await field1.field.update({ Title: "My Text"});

          let columnGroup = 'Socialiis';
          let fieldSchema = '<Field Type="Text" DisplayName="profilePic" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" ID="{571ed868-4226-472b-bc34-d783b00d8931}" SourceID="{60fda9ed-9447-4d2f-91fb-2d6b7eadd064}" StaticName="profilePic" Name="profilePic" ColName="nvarchar5" RowOrdinal="0" CustomFormatter="" Version="1"><Default>myDefaultValue</Default></Field>';
          const fieldXX: IFieldAddResult = await ensureResult.list.fields.createFieldAsXml(fieldSchema);


          const field2: IFieldAddResult = await ensureResult.list.fields.addText("keywords", 255, { Group: columnGroup });
          const field3: IFieldAddResult = await ensureResult.list.fields.addText("profilePic", 255, { Group: columnGroup });
          const field4: IFieldAddResult = await ensureResult.list.fields.addNumber("order", 0, 99, { Group: columnGroup, DefaultFormula: "99" });
          const field5: IFieldAddResult = await ensureResult.list.fields.addText("NavTitle", 255, { Group: columnGroup });

          const choices = ['blog','facebook','feed','github','home','instagram','linkedIn','location','office365-SPList','office365-SPPage','office365-SPSite','office365-team','office365-user','office365-YammerGroup','office365-YammerUser','office365-YammerSearch','stackExchange','stock','twitter','website','wikipedia','youtube-user','youtube-playlist','youtube-channel','youtube-video'];
          const field6: IFieldAddResult = await ensureResult.list.fields.addChoice("mediaObject", choices, ChoiceFieldFormatType.Dropdown, false, { Group: columnGroup });

          const field7: IFieldAddResult = await ensureResult.list.fields.addText("objectID", 255, { Group: columnGroup });
          const field8: IFieldAddResult = await ensureResult.list.fields.addText("url", 255, { Group: columnGroup });
          const field20 = await ensureResult.list.fields.addCalculated("mediaSource", 
          '=IF(ISNUMBER(FIND("-",mediaObject)),TRIM(LEFT(mediaObject,FIND("-",mediaObject)-1)),TRIM(mediaObject))', 
          DateTimeFieldFormatType.DateTime, FieldTypes.Text, { Group: columnGroup });

          const field21 = await ensureResult.list.fields.addCalculated("objectType", 
          '=IF(ISNUMBER(FIND("-",mediaObject)),TRIM(MID(mediaObject,FIND("-",mediaObject)+1,100)),"")', 
          DateTimeFieldFormatType.DateTime, FieldTypes.Text, { Group: columnGroup });

          let viewXml = '<View Name="{77880F39-3182-4CFF-8750-FA9817046AC5}" DefaultView="TRUE" MobileView="TRUE" MobileDefaultView="TRUE" Type="HTML" DisplayName="All Items" Url="/sites/Templates/Socialiis/Lists/EntityList/AllItems.aspx" Level="1" BaseViewID="1" ContentTypeID="0x" ImageUrl="/_layouts/15/images/generic.png?rev=47"><ViewFields><FieldRef Name="LinkTitle" /><FieldRef Name="keywords" /><FieldRef Name="profilePic" /><FieldRef Name="order0" /><FieldRef Name="NavTitle" /><FieldRef Name="mediaObject" /><FieldRef Name="objectID" /><FieldRef Name="url" /></ViewFields><ViewData /><Query><OrderBy><FieldRef Name="Title" /><FieldRef Name="order0" /></OrderBy></Query><Aggregations Value="Off" /><RowLimit Paged="TRUE">30</RowLimit><Mobile MobileItemLimit="3" MobileSimpleViewField="LinkTitle" /><XslLink Default="TRUE">main.xsl</XslLink><JSLink>clienttemplates.js</JSLink><Toolbar Type="Standard" /><ParameterBindings><ParameterBinding Name="NoAnnouncements" Location="Resource(wss,noXinviewofY_LIST)" /><ParameterBinding Name="NoAnnouncementsHowTo" Location="Resource(wss,noXinviewofY_DEFAULT)" /></ParameterBindings></View>';
          await ensureResult.list.views.getByTitle('All Items').setViewXml(viewXml);

          const resultVx = await ensureResult.list.views.add("My New View");
          viewXml = '<View Name="{B76BE63F-388D-402C-8B73-5405C5AFE019}" Type="HTML" DisplayName="Check Media Object" Url="/sites/Templates/Socialiis/Lists/EntityList/Check Media Object.aspx" Level="1" BaseViewID="1" ContentTypeID="0x" ImageUrl="/_layouts/15/images/generic.png?rev=47"><ViewFields><FieldRef Name="LinkTitle" /><FieldRef Name="keywords" /><FieldRef Name="profilePic" /><FieldRef Name="mediaObject" /><FieldRef Name="mediaSource" /><FieldRef Name="objectType" /></ViewFields><ViewData /><Query><OrderBy><FieldRef Name="Title" /><FieldRef Name="order0" /></OrderBy></Query><Aggregations Value="Off" /><RowLimit Paged="TRUE">30</RowLimit><Mobile MobileItemLimit="3" MobileSimpleViewField="LinkTitle" /><XslLink Default="TRUE">main.xsl</XslLink><JSLink>clienttemplates.js</JSLink><Toolbar Type="Standard" /><ParameterBindings><ParameterBinding Name="NoAnnouncements" Location="Resource(wss,noXinviewofY_LIST)" /><ParameterBinding Name="NoAnnouncementsHowTo" Location="Resource(wss,noXinviewofY_DEFAULT)" /></ParameterBindings></View>';
          await resultVx.view.setViewXml(viewXml);

          // the list is ready to be used
          result = true;
          alert(`Hey there!  Your ${myListName} list is all ready to go!`);
        } else {
          // the list already exists, double check the fields objectID
          try {
            const field2 = await ensureResult.list.fields.getByInternalNameOrTitle("keywords").get();
            const field3 = await ensureResult.list.fields.getByInternalNameOrTitle("profilePic").get();
            const field4 = await ensureResult.list.fields.getByInternalNameOrTitle("order").get();
            const field5 = await ensureResult.list.fields.getByInternalNameOrTitle("NavTitle").get();
            const field6 = await ensureResult.list.fields.getByInternalNameOrTitle("mediaSource").get();
            const field7 = await ensureResult.list.fields.getByInternalNameOrTitle("objectID").get();
            const field8 = await ensureResult.list.fields.getByInternalNameOrTitle("url").get();
            //const field9 = await ensureResult.list.fields.getByInternalNameOrTitle("PnPRedirectionEnabled").get();
            //const field10 = await ensureResult.list.fields.getByInternalNameOrTitle("PnPRedirectionEnabled").get();
            //const field11 = await ensureResult.list.fields.getByInternalNameOrTitle("PnPRedirectionEnabled").get();
            const field20 = await ensureResult.list.fields.getByInternalNameOrTitle("mediaSource").get();
            const field21 = await ensureResult.list.fields.getByInternalNameOrTitle("objectType").get();

            // if it is all good, then the list is ready to be used
            result = true;
            console.log(`Your ${myListName} list is already set up!`);
            alert(`Your ${myListName} list is already set up!`);
          } catch (e) {
            // if any of the fields does not exist, raise an exception in the console log
            let errMessage = this.getHelpfullError(e);
            console.log(`The ${myListName} list had this error:`, errMessage);

          }
        }
      }
    } catch (e) {
      // if we fail to create the list, raise an exception in the console log
      console.log(`Failed to create custom list ${myListName}.`, e, e.error);
    }

    return(result);
  }


  public getHelpfullError(e){
    let result = 'e';
    let errObj: {} = null;
      if (e.message) {
        let loc1 = e.message.indexOf("{\"");
        if (loc1 > 0) {
          result = e.message.substring(loc1);
          errObj = JSON.parse(result);
        }
    }
    result = errObj['odata.error']['message']['value'];
    console.log('errObj:',errObj);
    console.log('result:',result);
    return result;
  }
  

  private CreateThisList(oldVal: any): any {   
    let listName = this.properties.localListName ? this.properties.localListName : 'TestList4';
    let listDesc = 'Hey, this may actually work!';
    console.log('CreateThisList: oldVal', oldVal);
    let listCreated = this.ensureSocialiis7List(listName, listDesc);
    if ( listCreated ) { 
      this.properties.localListName= listName;
      this.properties.localListConfirmed= true;
      
    }
     return "test";  
  } 

  public render(): void {
    const element: React.ReactElement<IAssetBuilderProps > = React.createElement(
      AssetBuilder,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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

                PropertyPaneButton('ClickMe',  
                {  
                 text: "Create/Verify List",  
                 buttonType: PropertyPaneButtonType.Normal,  
                 onClick: this.CreateThisList.bind(this)  
                }), 


                
                PropertyPaneLabel('confirmation', {
                  text: this.properties.localListConfirmed ? this.properties.localListName + ' List is available' : 'Verify or Create your list!'
                }),

                PropertyPaneTextField('localListName', {
                  label: strings.LocalListFieldLabel
                }),

                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {  
    if (propertyPath === 'localListName' &&  newValue) {  
      this.properties.localListName=newValue;  
    }  
  } 

}
