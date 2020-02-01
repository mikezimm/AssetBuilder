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

const createThisList = 'TestList'

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
        "Redirections",
        100,
        true);

      // if we've got the list
      if (ensureResult.list != null) {
        // if the list has just been created
        if (ensureResult.created) {
          // we need to add the custom fields to the list
          //https://pnp.github.io/pnpjs/sp/lists/#ensure-that-a-list-exists-by-title
          //https://pnp.github.io/pnpjs/sp/fields/


          const choices = [`ChoiceA`, `ChoiceB`, `ChoiceC`];

          const field1: IFieldAddResult = await ensureResult.list.fields.addText("MyText", 255, { Group: "My Group" });
          await field1.field.update({ Title: "My Text"});

          const field2: IFieldAddResult = await ensureResult.list.fields.addChoice("MyChoice", choices, ChoiceFieldFormatType.Dropdown, false, { Group: "My Group" });
          await field2.field.update({ Title: "My Choices"});

          const field3 = await ensureResult.list.fields.addCalculated("MyCalculation", "=Modified+1", DateTimeFieldFormatType.DateOnly, FieldTypes.DateTime, { Group: "MyGroup" });
          await field3.field.update({ Title: "My Calculation"});

          const sourceUrlFieldAddResult: IFieldAddResult = await ensureResult.list.fields.addUrl(
            "PnPSourceUrl", UrlFieldFormatType.Hyperlink,
            { Required: true });
          await sourceUrlFieldAddResult.field.update({ Title: "Source URL"});
          const destinationUrlFieldAddResult: IFieldAddResult = await ensureResult.list.fields.addUrl(
            "PnPDestinationUrl", UrlFieldFormatType.Hyperlink,
            { Required: true });
          await destinationUrlFieldAddResult.field.update({ Title: "Destination URL"});
          const redirectionEnabledFieldAddResult: IFieldAddResult = await ensureResult.list.fields.addBoolean(
            "PnPRedirectionEnabled",
            { Required: true });
          await redirectionEnabledFieldAddResult.field.update({ Title: "Redirection Enabled"});

          // the list is ready to be used
          result = true;
        } else {
          // the list already exists, double check the fields
          try {
            const sourceUrlField = await ensureResult.list.fields.getByInternalNameOrTitle("PnPSourceUrl").get();
            const destinationUrlField = await ensureResult.list.fields.getByInternalNameOrTitle("PnPDestinationUrl").get();
            const redirectionEnabledField = await ensureResult.list.fields.getByInternalNameOrTitle("PnPRedirectionEnabled").get();

            // if it is all good, then the list is ready to be used
            result = true;
          } catch (e) {
            // if any of the fields does not exist, raise an exception in the console log
            console.log(`The ${createThisList} list does not match the expected fields definition.`);
          }
        }
      }
    } catch (e) {
      // if we fail to create the list, raise an exception in the console log
      console.log(`Failed to create custom list ${createThisList}.`);
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
