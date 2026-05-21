import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'EmployeeHubShoutOutAcknowledgeWebPartStrings';
import EmployeeHubShoutOutAcknowledge from './components/EmployeeHubShoutOutAcknowledge';
import { IEmployeeHubShoutOutAcknowledgeProps } from './components/IEmployeeHubShoutOutAcknowledgeProps';
import { sp } from '@pnp/sp/presets/all';

export interface IEmployeeHubShoutOutAcknowledgeWebPartProps {
  description: string;
  ShoutOutTitle: any;
  ShoutOutDescription: any;
  Link: any;
}

export default class EmployeeHubShoutOutAcknowledgeWebPart extends BaseClientSideWebPart<IEmployeeHubShoutOutAcknowledgeWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    sp.setup({
      spfxContext: this.context
    });

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IEmployeeHubShoutOutAcknowledgeProps> = React.createElement(
      EmployeeHubShoutOutAcknowledge,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        ShoutOutTitle: this.properties.ShoutOutTitle,
        ShoutOutDescription: this.properties.ShoutOutDescription,
        Link: this.properties.Link
      }
    );

    ReactDom.render(element, this.domElement);
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
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

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
                PropertyPaneTextField('ShoutOutTitle', {
                  label: "ShoutOut Title"
                }),
                PropertyPaneTextField('ShoutOutDescription', {
                  label: "ShoutOut Description"
                }),
                PropertyPaneTextField('Link', {
                  label: "Link"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
