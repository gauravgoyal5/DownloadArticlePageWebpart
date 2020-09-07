import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import {
  BaseClientSideWebPart,
  WebPartContext,
} from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";

import * as strings from "DownloadArticlePageWebPartStrings";
import DownloadArticlePage from "./components/DownloadArticlePage";
import { IDownloadArticlePageProps } from "./components/IDownloadArticlePageProps";

export interface IDownloadArticlePageWebPartProps {
  description: string;
  context: WebPartContext;
}

export default class DownloadArticlePageWebPart extends BaseClientSideWebPart<
  IDownloadArticlePageWebPartProps
> {
  protected onInit() {
    // initialize the PnP Js
    sp.setup({
      spfxContext: this.context,
    });

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IDownloadArticlePageProps> = React.createElement(
      DownloadArticlePage,
      {
        description: this.properties.description,
        context: this.context,
      }
    );

    ReactDom.render(element, this.domElement);
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
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
