import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'PnPHttpClientWebPartStrings';
import PnPHttpClient from './components/PnPHttpClient';
import { IPnPHttpClientProps } from './components/IPnPHttpClientProps';

import { ICountryListItem } from '../../models';
import { getSP } from './pnpjsConfig';
import { SPFI } from '@pnp/sp';

export interface IPnPHttpClientWebPartProps {
  description: string;
}

export default class PnPHttpClientWebPart extends BaseClientSideWebPart<IPnPHttpClientWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _countries: ICountryListItem[] = [];
  private _sp: SPFI;

  public render(): void {
    const element: React.ReactElement<IPnPHttpClientProps> = React.createElement(
      PnPHttpClient,
      {
        spListItems: this._countries,
        onGetListItems: this._onGetListItems,
        onAddListItem: this._onAddListItem,
        onUpdateListItem: this._onUpdateListItem,
        onDeleteListItem: this._onDeleteListItem,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _onGetListItems = async (): Promise<void> => {
    const response: ICountryListItem[] = await this._getListItems();
    this._countries = response;
    this.render();
  }

  private async _getListItems(): Promise<ICountryListItem[]> {
    var getItems = await this._sp.web.lists.getByTitle('Countries').items();
    console.log(getItems);
  
    return getItems as ICountryListItem[];
  }

  private _onAddListItem = async (): Promise<void> => {
    await this._sp.web.lists.getByTitle('Countries').items.add({Title:"New Pnp Item Created"});
  
    const getResponse: ICountryListItem[] = await this._getListItems();
    this._countries = getResponse;
    this.render();
  }
  
  private _onUpdateListItem = async (): Promise<void> => {
    var item = await this._sp.web.lists.getByTitle("Countries").items.select("ID").orderBy("ID",false).top(1)();
    var itemId: number = item[0].Id;
    await this._sp.web.lists.getByTitle("Countries").items.getById(itemId).update({Title:"Pnp Item Updated"});

    const getResponse: ICountryListItem[] = await this._getListItems();
    this._countries = getResponse;
    this.render();
  }
  
  private _onDeleteListItem = async (): Promise<void> => {
    var item = await this._sp.web.lists.getByTitle('Countries').items.select("ID").orderBy("ID",false).top(1)();
    var itemId: number = item[0].Id;
    await this._sp.web.lists.getByTitle('Countries').items.getById(itemId).delete();
  
    const getResponse: ICountryListItem[] = await this._getListItems();
    this._countries = getResponse;
    this.render();
  }

  protected onInit(): Promise<void> {
    this._sp = getSP(this.context);
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
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
