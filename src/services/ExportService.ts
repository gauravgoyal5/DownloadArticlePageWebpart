import { SPService } from "./SPService";
import { PdfService } from "./PdfService";
import { HttpClientResponse } from "@microsoft/sp-http";
import { Item } from "@pnp/sp/items";

export class ExportService {
  private listItem: any;
  private listTitle: string;
  private pdfService: PdfService;

  constructor(listTitle: string, item: any) {
    this.listItem = item;
    this.listTitle = listTitle;
    this.pdfService = new PdfService();
  }

  /**
   * Export the article to PDF by extracting article information as text
   */
  public exportItem() {
    const spService = new SPService(this.listTitle, this.listItem);

    return new Promise((resolve, reject) => {
      spService.getItem<any>().then((item) => {
        console.log(item);

        this.pdfService.exportItemToPdf(item);
        resolve();
      });
    });
  }
}
