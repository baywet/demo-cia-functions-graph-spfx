import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";

import * as strings from "NiceWebPartStrings";
import Nice from "./components/Nice";
import { INiceProps } from "./components/INiceProps";
import { NiceService } from "./data";

export interface INiceWebPartProps {
  description: string;
}

export default class NiceWebPart extends BaseClientSideWebPart<INiceWebPartProps> {
  private niceService: NiceService;
  public async render(): Promise<void> {
    this.niceService = new NiceService(this.context.serviceScope);
    const Score: number = (await this.niceService.GetDataForCurrentUser()).Score;
    const element: React.ReactElement<INiceProps > = React.createElement(
      Nice,
      {
        description: this.properties.description,
        score: Score,
        registerWebHook: async () => await this.niceService.RegisterWebHook()
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
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
                PropertyPaneTextField("description", {
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
