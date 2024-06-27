import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  // PropertyPaneLabel,
  // PropertyPaneLink,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneDropdown,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "SpfxWebPartStrings";
import Spfx from "./components/Spfx";
import { ISpfxProps } from "./components/ISpfxProps";

import { escape } from "@microsoft/sp-lodash-subset";
import { SPHttpClient } from "@microsoft/sp-http";

export interface ISpfxWebPartProps {
  description: string;
  multi: string;
  check: boolean;
  toggle: boolean;
  dropDown: string;
  slider: string;
  List: string;
}

export default class SpfxWebPart extends BaseClientSideWebPart<ISpfxWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  public render(): void {
    const element: React.ReactElement<ISpfxProps> = React.createElement(Spfx, {
      description: this.properties.description,
      isDarkTheme: this._isDarkTheme,
      environmentMessage: this._environmentMessage,
      hasTeamsContext: !!this.context.sdks.microsoftTeams,
      userDisplayName: this.context.pageContext.user.displayName,
    });

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then((message) => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
            case "TeamsModern":
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                  onGetErrorMessage: this.validateDescription.bind(this),
                }),
                PropertyPaneTextField("multi", {
                  label: "Enter your comments here!",
                  multiline: true,
                }),
                PropertyPaneCheckbox("check", {
                  text: "Checkbox",
                }),

                PropertyPaneToggle("toggle", {
                  label: "Dark",
                  onText: "On",
                  offText: "Off",
                }),
                PropertyPaneDropdown("dropDown", {
                  label: "Framework",
                  options: [
                    { key: "1", text: "React" },
                    { key: "2", text: "Angular" },
                    { key: "3", text: "Vuejs" },
                    { key: "4", text: "HandleBar" },
                  ],
                }),
                PropertyPaneSlider("slider", {
                  label: "Rate your scale",
                  min: 1,
                  max: 5,
                }),
                PropertyPaneTextField("List,", {
                  label: "List Name",
                  onGetErrorMessage: this.validateList.bind(this),
                  deferredValidationTime: 500,
                }),
              ],
            },
          ],
        },
      ],
    };
  }

  private validateDescription(value: string): string {
    if (value === null || value.trim().length === 0) {
      return "Provide a description";
    }

    if (value.length > 40) {
      return "Description should not be longer than 40 characters";
    }

    return "";
  }

  private async validateList(value: string) {
    if (value === null || value.trim().length === 0) {
      return "Provide the List name";
    }

    try {
      let response = await this.context.spHttpClient.get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists/getByTitle('${escape(value)}')?$select=Id`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) return "";
      else if (response.status === 404) {
        return ` List ${escape(value)} doesn't exist in the current site`;
      } else {
        return `Error: ${response.statusText}. Please try again`;
      }
    } catch (error) {
      return error.message;
    }
  }

  protected get disableReactivePropertyChanges(): boolean {
    // True for non reactive web part changes
    return false;
  }
}
