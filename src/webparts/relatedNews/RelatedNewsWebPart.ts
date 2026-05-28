import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';

import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'RelatedNewsWebPartStrings';

import RelatedNews from './components/RelatedNews';
import { IRelatedNewsProps } from './components/IRelatedNewsProps';

import { sp } from '@pnp/sp/presets/all';

export interface IRelatedNewsWebPartProps {
  description: string;
  category: string;
}

export default class RelatedNewsWebPart extends BaseClientSideWebPart<IRelatedNewsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  // Dynamic dropdown options
  private _categoryOptions: IPropertyPaneDropdownOption[] = [];

  protected async onInit(): Promise<void> {

    this._environmentMessage = this._getEnvironmentMessage();

    sp.setup({
      spfxContext: this.context
    });

    // Load choice field values
    await this.loadCategoryOptions();

    return super.onInit();
  }

  // Load SharePoint Choice Field Options
  private async loadCategoryOptions(): Promise<void> {

    try {

      const field: any = await sp.web.lists
        .getByTitle("News Announcements") // SharePoint List Name
        .fields
        .getByInternalNameOrTitle("NewsCategory")(); // Choice Column Internal Name

      this._categoryOptions = field.Choices.map((choice: string) => {
        return {
          key: choice,
          text: choice
        };
      });

      // Refresh Property Pane
      this.context.propertyPane.refresh();

    } catch (error) {

      console.log("Error loading category choices", error);

    }
  }

  public render(): void {

    const element: React.ReactElement<IRelatedNewsProps> = React.createElement(
      RelatedNews,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        category: this.properties.category
          ? this.properties.category
          : "Company"
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _getEnvironmentMessage(): string {

    if (!!this.context.sdks.microsoftTeams) {

      return this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentTeams
        : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost
      ? strings.AppLocalEnvironmentSharePoint
      : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {

    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;

    const { semanticColors } = currentTheme;

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
          groups: [
            {
              groupName: "Filter Options",
              groupFields: [

                PropertyPaneDropdown('category', {
                  label: 'Select Category',
                  options: this._categoryOptions
                })

              ]
            }
          ]
        }
      ]
    };
  }
}