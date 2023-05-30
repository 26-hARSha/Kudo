import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "KudoWebPartStrings";
import Kudo from "./components/Kudo";
import { IKudoProps } from "./components/IKudoProps";

export interface IKudoWebPartProps {
  seeAllPageURL: string;
  emptyMessage: string;
  noofKudos: string;
  errorMessage: string;
  componentTitle: string;
  listName: string;
  webHeight: string;
  description: string;
}

export default class KudoWebPart extends BaseClientSideWebPart<IKudoWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  public render(): void {
    const element: React.ReactElement<IKudoProps> = React.createElement(Kudo, {
      description: this.properties.description,
      isDarkTheme: this._isDarkTheme,
      environmentMessage: this._environmentMessage,
      hasTeamsContext: !!this.context.sdks.microsoftTeams,
      userDisplayName: this.context.pageContext.user.displayName,
      siteURL: this.context.pageContext.web.absoluteUrl,
      context: this.context,
      listName: this.properties.listName,
      componentTitle: this.properties.componentTitle,
      webHeight: this.properties.webHeight,
      emptyMessage: this.properties.emptyMessage,
      seeAllPageURL: this.properties.seeAllPageURL,
      noofKudos: this.properties.noofKudos,
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
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error("Unknown host");
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
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected onAfterPropertyPaneChangesApplied(): void {
    ReactDom.unmountComponentAtNode(this.domElement);

    this.render();
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
          groups: [
            {
              groupFields: [
                PropertyPaneTextField("componentTitle", {
                  label: "Component Title",
                }),
                PropertyPaneTextField("listName", {
                  label: "Enter List Name",
                }),
                PropertyPaneTextField("emptyMessage", {
                  label: "Text, if no Info available",
                }),
                PropertyPaneTextField("noofKudos", {
                  label: "No. of Kudos to Show",
                }),
                PropertyPaneTextField("seeAllPageURL", {
                  label: "See All Page URL",
                }),
                PropertyPaneSlider("webHeight", {
                  label: "Compnent Height",
                  min: 300,
                  max: 1000,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
