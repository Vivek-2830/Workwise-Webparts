import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'EmployeehubPerkofMonthWebPartStrings';
import EmployeehubPerkofMonth from './components/EmployeehubPerkofMonth';
import { IEmployeehubPerkofMonthProps } from './components/IEmployeehubPerkofMonthProps';
import { sp } from '@pnp/sp/presets/all';
import { IFilePickerResult, PropertyFieldFilePicker } from '@pnp/spfx-property-controls/lib/propertyFields/filePicker';

export interface IEmployeehubPerkofMonthWebPartProps {
  description: string;
  PerkMonthImage: IFilePickerResult;
  PerkMonthDescription: string;
  LinkButton: string;
}

export default class EmployeehubPerkofMonthWebPart extends BaseClientSideWebPart<IEmployeehubPerkofMonthWebPartProps> {

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
    const element: React.ReactElement<IEmployeehubPerkofMonthProps> = React.createElement(
      EmployeehubPerkofMonth,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        PerkMonthImage: this.properties.PerkMonthImage,
        PerkMonthDescription: this.properties.PerkMonthDescription,
        LinkButton: this.properties.LinkButton
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
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('LinkButton', {
                  label: 'Add PerkMonth Link',
                }),
                PropertyFieldFilePicker("PerkMonth Image", {
                  context: this.context,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: (e: IFilePickerResult) => {
                    console.log(e);
                    this.properties.PerkMonthImage = e;
                  },
                  onChanged: (e: IFilePickerResult) => {
                    console.log(e);
                    this.properties.PerkMonthImage = e;
                  },
                  buttonLabel: "Upload PerkMonth Image",
                  label: "Our PerkMonth  Image",
                  key: "FilePickerID",
                  filePickerResult: this.properties.PerkMonthImage,
                  hideLocalUploadTab: true,
                }),
                PropertyPaneTextField('PerkMonthDescription', {
                  label: "PerkMonth Description"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
