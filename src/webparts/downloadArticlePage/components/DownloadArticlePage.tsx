import * as React from "react";
import styles from "./DownloadArticlePage.module.scss";
import { IDownloadArticlePageProps } from "./IDownloadArticlePageProps";
import { escape } from "@microsoft/sp-lodash-subset";
import {
  SPHttpClient,
  SPHttpClientResponse,
  IHttpClientOptions,
  HttpClientResponse,
} from "@microsoft/sp-http";

import { ExportService } from "../../../services/ExportService";

export default class DownloadArticlePage extends React.Component<
  IDownloadArticlePageProps,
  {}
> {
  private _spHttpClient: SPHttpClient;
  private _currentWebUrl: string;

  constructor(props: IDownloadArticlePageProps) {
    super(props);
    this._spHttpClient = props.context.spHttpClient;
    this._currentWebUrl = props.context.pageContext.web.absoluteUrl;
  }

  public render(): React.ReactElement<IDownloadArticlePageProps> {
    const exportService = new ExportService(
      "Site Pages",
      this.props.context.pageContext.listItem
    );
    return (
      <div className={styles.downloadArticlePage}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <button
                id="downloadButton"
                onClick={() => exportService.exportItem()}
              >
                Download
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
