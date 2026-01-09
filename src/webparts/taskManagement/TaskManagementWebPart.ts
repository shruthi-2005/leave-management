import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'TaskManagementWebPartStrings';
import TaskManagement from './components/TaskManagement';
import { ITaskManagementProps } from './components/ITaskManagementProps';

export interface ITaskManagementWebPartProps {
  description: string;
}

export default class TaskManagementWebPart extends BaseClientSideWebPart<ITaskManagementWebPartProps> {

  public render(): void {
    // ✅ Main entry point: detect Teams or browser context
    if (this.context.sdks?.microsoftTeams) {
      this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then((teamsContext: any) => {
          console.log('🔍 Teams context:', teamsContext);

          if (teamsContext?.subEntityId) {
            // Directly from Teams SDK
            const taskId = teamsContext.subEntityId.replace('TASK_', '');
            console.log('🔗 Loaded from Teams SDK subEntityId:', taskId);
            this.renderReact(taskId);
          } else {
            console.log('ℹ️ No subEntityId in Teams SDK, checking URL...');
            this.handleBrowserDeeplink();
          }
        })
        .catch((err: any) => {
          console.error('❌ Error getting Teams SDK context:', err);
          this.handleBrowserDeeplink();
        });
    } else {
      // ✅ Running in browser (SharePoint page or Teams launcher link)
      this.handleBrowserDeeplink();
    }
  }

  // 🔹 Handle taskId from browser URL or nested Teams launcher link
  private handleBrowserDeeplink(): void {
    const urlParams = new URLSearchParams(window.location.search);
    let taskId: string | null = null;

    // 1️⃣ Standard browser param
    if (urlParams.has('taskId')) {
      taskId = urlParams.get('taskId');
    }

    // 2️⃣ If not found, check for nested Teams launcher JSON
    if (!taskId && urlParams.has('url')) {
      try {
        const nestedUrl = decodeURIComponent(urlParams.get('url') || '');
        const nestedParams = new URLSearchParams(nestedUrl.split('?')[1]);
        if (nestedParams.has('context')) {
          const contextObj = JSON.parse(decodeURIComponent(nestedParams.get('context')!));
          if (contextObj.subEntityId) {
            taskId = contextObj.subEntityId.replace('TASK_', '');
            console.log('🔗 Loaded from Teams launcher URL:', taskId);
          }
        }
      } catch (err) {
        console.error('❌ Failed to parse nested Teams URL context:', err);
      }
    }

    console.log('🌐 Loaded from browser URL:', taskId);
    this.renderReact(taskId);
  }

  // 🔹 Render the React component
  private renderReact(taskId: string | null): void {
    const element: React.ReactElement<ITaskManagementProps> = React.createElement(
      TaskManagement,
      {
        context: this.context,
        taskId: taskId
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      // optional: use environment message if needed
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }
          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;

    const { semanticColors } = currentTheme;
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
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', { label: strings.DescriptionFieldLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}