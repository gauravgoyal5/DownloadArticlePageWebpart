import * as jsPDF from "jspdf";
// import * as moment from "moment";
import "jspdf-autotable";
import { DateTimeFieldFormatType } from "@pnp/sp/fields";
import { stringIsNullOrEmpty } from "@pnp/common";
import html2canvas from "html2canvas";
import { PdfConstants } from "./PdfConstants";

interface Field {
  label: string;
  value: string;
}

export class PdfService {
  constructor() {}

  /**
   * Creates and saves the PDF document for the article page.
   * @param item page item.
   */
  public exportItemToPdf(item: any) {
    const doc = this.generatePdf(item.Title, this.getFields(item));

    doc.save(`${item.Title}.pdf`);
  }

  /**
   * Export the article to PDF by taking the article content as image
   * @param canvasContent canvas content
   * @param item page list item
   */
  public exportItemToPdfAsImage(canvasContent: HTMLElement, item: any) {
    const myinput = canvasContent; //document.getElementById("spPageCanvasContent");
    html2canvas(myinput).then((canvas) => {
      var imgWidth = 200;
      var pageHeight = 290;
      var imgHeight = (canvas.height * imgWidth) / canvas.width;
      var heightLeft = imgHeight;
      const imgData = canvas.toDataURL("image/png");
      const mynewpdf = new jsPDF("p", "mm", "a4");
      var position = 0;
      mynewpdf.addImage(imgData, "JPEG", 5, position, imgWidth, imgHeight);
      mynewpdf.save(`${item.Title}.pdf`);
    });
  }

  /**
   * @param item article page item.
   */
  private getFields(item: any): Field[] {
    return [
      { label: "PublishedDate", value: item.FirstPublishedDate },
      { label: "BannerUrl", value: item.BannerImageUrl.Url },
      { label: "CanvasContent1", value: item.CanvasContent1 },
    ];
  }

  /**
   * Returns the generated PDF document that contains a header, page image and article body.
   * @param title Header for the file.
   * @param fields List of fields to process.
   */
  private generatePdf(title: string, fields: Field[]): jsPDF {
    let start = 40;

    const doc = new jsPDF();
    var elementHandler = {
      "#ignorePDF": (element, renderer) => {
        return true;
      },
    };

    // Render a header
    doc.setFontSize(PdfConstants.h1FontSize);
    doc.setFontType("bold");
    doc.text(PdfConstants.offset, PdfConstants.offset, title);

    doc.setFontSize(PdfConstants.textFontSize);
    doc.setFontType("normal");

    // Render a page fields
    fields.forEach((field, index) => {
      switch (field.label) {
        case "BannerUrl": {
          //Add image
          start += PdfConstants.margin * index;
          if (!stringIsNullOrEmpty(field.value)) {
            var logo_url = field.value;
            var imgData = this.imageToBase64(logo_url);
            doc.addImage(
              imgData,
              "JPEG",
              PdfConstants.offset,
              start,
              PdfConstants.imgWidth,
              PdfConstants.imgHeight
            );
          }

          break;
        }
        case "PublishedDate": {
          //Add publish date
          start += PdfConstants.margin * index;
          var publishDate = new Date(field.value);

          doc.text(PdfConstants.offset, start, "Published Date:");
          doc.text(
            PdfConstants.dateOffset,
            start,
            publishDate.toLocaleDateString("en-US")
          );

          break;
        }
        case "CanvasContent1": {
          //Add article content
          start += PdfConstants.margin * index + PdfConstants.imgHeight;
          let content = this.formatContent(field.value);

          doc.fromHTML(content, PdfConstants.offset, start, {
            width: 180,
            elementHandlers: elementHandler,
          });
          break;
        }
      }
    });

    return doc;
  }
  /**
   * This removed unwanted details like GUIDs of the webpart added on the page using Regex match and leaves only the article content
   * @param canvasContent article main content
   */
  private formatContent(canvasContent: string) {
    var content = canvasContent;
    var formattedContent = "";

    var matches = content.match(PdfConstants.primaryRegex); //"<div[^>]*?data-sp-rte[^>]*?>(.*?)</div>"

    if (matches)
      var secondaryMatch = matches[1].match(PdfConstants.secondaryRejex);
    else formattedContent = matches[1];

    if (secondaryMatch) {
      for (let i = 0; i < secondaryMatch.length; i++) {
        formattedContent = matches[1].replace(secondaryMatch[i], "");
      }
    }

    return formattedContent;
  }

  /**
   * Takes the url of the banner image and generates Base64 representation of it to be used in writing to the PDF
   * @param URL image url
   */
  private imageToBase64(URL: string) {
    let image;
    image = new Image();
    image.crossOrigin = "Anonymous";
    image.addEventListener("load", () => {
      let canvas = document.createElement("canvas");
      let context = canvas.getContext("2d");
      canvas.width = image.width;
      canvas.height = image.height;
      context.drawImage(image, 0, 0);
      try {
        localStorage.setItem(
          "saved-image-example",
          canvas.toDataURL("image/png")
        );
      } catch (err) {
        console.error(err);
      }
    });

    image.src = URL;
    return image;
  }
}
