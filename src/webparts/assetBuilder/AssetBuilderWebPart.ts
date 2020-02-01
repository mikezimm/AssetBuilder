import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
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

export interface IAssetBuilderWebPartProps {
  description: string;
}

const LOG_SOURCE: string = 'RedirectApplicationCustomizer';

const createThisList = 'TestList2'

export default class AssetBuilderWebPart extends BaseClientSideWebPart<IAssetBuilderWebPartProps> {

  @override
  public async onInit(): Promise<void> {

    // initialize PnP JS library to play with SPFx contenxt
    sp.setup({
      spfxContext: this.context
    });

    // read the server relative URL of the current page from Legacy Page Context
    const currentPageRelativeUrl: string = this.context.pageContext.legacyPageContext.serverRequestPath;

    if (await this.ensureRedirectionsList(createThisList)) {
      alert('Done Ensuring!');
    }

  }

  // this method ensures that the Redirections lists exists, or if it doesn't exist
  // it creates it, as long as the currently connected user has proper permissions
  private async ensureRedirectionsList(createThisList: string): Promise<boolean> {

    let result: boolean = false;

    try {
      const ensureResult = await sp.web.lists.ensure(createThisList,
        "My List Description",
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


          /* Url Field Sample
          const sourceUrlFieldAddResult: IFieldAddResult = await ensureResult.list.fields.addUrl(
            "PnPSourceUrl", UrlFieldFormatType.Hyperlink,
            { Required: true });
          await sourceUrlFieldAddResult.field.update({ Title: "Source URL"});
          */

          /* Boolean Field Sample
          const redirectionEnabledFieldAddResult: IFieldAddResult = await ensureResult.list.fields.addBoolean(
            "PnPRedirectionEnabled",
            { Required: true });
          await redirectionEnabledFieldAddResult.field.update({ Title: "Redirection Enabled"});
          */

          // the list is ready to be used
          result = true;
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
          } catch (e) {
            // if any of the fields does not exist, raise an exception in the console log
            console.log(`The ${createThisList} list does not match the expected fields definition.`, e, e.odata.error.message);
          }
        }
      }
    } catch (e) {
      // if we fail to create the list, raise an exception in the console log
      console.log(`Failed to create custom list ${createThisList}.`, e, e.error);
    }

    return(result);
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
}
