
           import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import TripShare from './components/TripShare';
import { ITripShareProps } from './components/ITripShareProps';

export interface ITripShareWebPartProps {
  description: string;
}

export default class TripShareWebPart extends BaseClientSideWebPart<ITripShareWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ITripShareProps> = React.createElement(
      TripShare,
      {
        context: this.context,
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();
    this._isDarkTheme = this.context.sdks.microsoftTeams
      ? this.context.sdks.microsoftTeams.context.theme === "dark"
      : false;

    return super.onInit();
  }

  private _getEnvironmentMessage(): string {
    if (this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost
        ? 'Local Environment (Teams)'
        : 'Teams Environment';
    }

    return this.context.isServedFromLocalhost
      ? 'Local Environment'
      : 'SharePoint Environment';
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
          header: { description: "Configure your Trip Expense app" },
          groups: [
            {
              groupName: "Settings",
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Web part description"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}