/*
 *  Power BI Visual CLI
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */
"use strict";

import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbiVisualsApi.extensibility.visual.IVisualHost;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstance = powerbi.VisualObjectInstance;
import DataView = powerbi.DataView;
import VisualObjectInstanceEnumerationObject = powerbi.VisualObjectInstanceEnumerationObject;

import { base64PdfData } from "./base64Pdf";
import { base64ExcelData} from "./base64Excel";
import { excelData } from "./sampleExcelData";
import { data } from "./samplePdfData";
import { VisualSettings } from "./settings";
import powerbiVisualsApi from "powerbi-visuals-api";
export class Visual implements IVisual {
  private target: HTMLElement;
  private updateCount: number;
  private settings: VisualSettings;
  private textNode: Text;
  private visualHost: IVisualHost;

  constructor(options: VisualConstructorOptions) {
    console.log("Visual constructor", options);
    this.target = options.element;
    this.visualHost = options.host;
    this.updateCount = 0;
    if (document) {
      const api_pdf_btn: HTMLElement = document.createElement("input");
      api_pdf_btn.setAttribute("type", "button");
      api_pdf_btn.setAttribute("value", "Download PDF Api");
      api_pdf_btn.onclick = () => {
        console.log("working...");
        this.visualHost.downloadService.exportVisualsContent(
            base64PdfData,
          "export.pdf",
          "base64",
          "test"
        );
      };
      this.target.appendChild(api_pdf_btn);

       const pdfBlob = this.b64toBlob(base64PdfData, "application/pdf");
      const pdf_url: string = window.URL.createObjectURL(pdfBlob);
       const pdf_btn: HTMLElement = document.createElement("a");
       pdf_btn.setAttribute("href", pdf_url);
       pdf_btn.innerHTML = "Download PDF sample";
       this.target.appendChild(pdf_btn);

      const br: HTMLElement = document.createElement("br");
      this.target.appendChild(br);

      const api_excel_btn: HTMLElement = document.createElement("input");
      api_excel_btn.setAttribute("type", "button");
      api_excel_btn.setAttribute("value", "Download Excel Api");
      api_excel_btn.onclick = () => {
        this.visualHost.downloadService.exportVisualsContent(
            base64ExcelData,
          "export.xlsx",
          "base64",
          "test"
        );
      };
      this.target.appendChild(api_excel_btn);

      // const excelBlob = this.b64toBlob(
      //   base64ExcelData,
      //   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      // );
      // const excel_url: string = window.URL.createObjectURL(excelBlob);
      // const excel_btn: HTMLElement = document.createElement("a");
      // excel_btn.setAttribute("href", excel_url);
      // excel_btn.innerHTML = "Download Excel sample";
      // this.target.appendChild(excel_btn);
    }
  }

  private b64toBlob(b64Data, contentType = "", sliceSize = 512) {
    const byteCharacters = atob(b64Data);
    const byteArrays = [];

    for (let offset = 0; offset < byteCharacters.length; offset += sliceSize) {
      const slice = byteCharacters.slice(offset, offset + sliceSize);

      const byteNumbers = new Array(slice.length);
      for (let i = 0; i < slice.length; i++) {
        byteNumbers[i] = slice.charCodeAt(i);
      }

      const byteArray = new Uint8Array(byteNumbers);
      byteArrays.push(byteArray);
    }

    const blob = new Blob(byteArrays, { type: contentType });
    return blob;
  }

  public update(options: VisualUpdateOptions) {
    this.settings = Visual.parseSettings(
      options && options.dataViews && options.dataViews[0]
    );
    console.log("Visual update", options);
    if (this.textNode) {
      this.textNode.textContent = (this.updateCount++).toString();
    }
  }

  private static parseSettings(dataView: DataView): VisualSettings {
    return <VisualSettings>VisualSettings.parse(dataView);
  }

  /**
   * This function gets called for each of the objects defined in the capabilities files and allows you to select which of the
   * objects and properties you want to expose to the users in the property pane.
   *
   */
  public enumerateObjectInstances(
    options: EnumerateVisualObjectInstancesOptions
  ): VisualObjectInstance[] | VisualObjectInstanceEnumerationObject {
    return VisualSettings.enumerateObjectInstances(
      this.settings || VisualSettings.getDefault(),
      options
    );
  }
}
