import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneButton,
  PropertyPaneLabel,
  PropertyPaneButtonType
} from '@microsoft/sp-property-pane';

import * as strings from 'SohoWikiVaultWebPartStrings';
import { ISohoWikiVaultProps, SohoWikiVault } from './components';
import { update, get } from '@microsoft/sp-lodash-subset';

import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';

export interface ISohoWikiVaultWebPartProps {
  webPartTitle: string;
  setupComplete: boolean;
  info: string;
}
// Data Provider
import { PropertyPaneDataProvider } from './dataProvider';
// Pnp
import { sp } from "@pnp/sp";

export default class SohoWikiVaultWebPart extends BaseClientSideWebPart<ISohoWikiVaultWebPartProps> {
  private propertyPaneData: PropertyPaneDataProvider;
  private loadingIndicator: boolean = false;
  private error: string = '';

  public async onInit(): Promise<void> {
    await super.onInit();
    sp.setup({
      spfxContext: this.context
    });
    this.propertyPaneData = new PropertyPaneDataProvider(this.context);
  }
  public render(): void {
    var queryParms = new UrlQueryParameterCollection(window.location.href);
    let wikiId: string = '';
    if (queryParms.getValue("wikiId")) {
      wikiId = queryParms.getValue("wikiId");
    }
    console.log(wikiId);
    const element: React.ReactElement<ISohoWikiVaultProps> = React.createElement(
      SohoWikiVault,
      {
        // webPartTitle: this.properties.webPartTitle,
        setupComelete: this.properties.setupComplete,
        context: this.context,
        wikiId: wikiId
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
  private ButtonClick = async (val: any) => {
    this.loadingIndicator = true;
    const updated: boolean = await this.propertyPaneData.createConfigLists();
    if (updated) {
      update(this.properties, 'setupComplete', (): any => { return true; });
      update(this.properties, 'info', (): any => { return 'All setups are Installed.'; });
    } else {
      this.error = 'You need full control on the site to complete the setup';
    }
    this.loadingIndicator = false;
    this.context.propertyPane.refresh();
    this.render();
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
                // PropertyPaneTextField('webPartTitle', {
                //   label: strings.webPartTitleFieldLabel
                // }),
                PropertyPaneButton('installSettings',
                  {
                    text: strings.installSettingsButtonLabel,
                    buttonType: PropertyPaneButtonType.Normal,
                    onClick: this.ButtonClick.bind(this),
                    disabled: this.properties.setupComplete
                  }),
                PropertyPaneLabel('', {
                  text: this.properties.info
                }),
                PropertyPaneLabel('', {
                  text: this.error,
                }),
              ]
            }
          ]
        }
      ],
      showLoadingIndicator: this.loadingIndicator
    };
  }
}
