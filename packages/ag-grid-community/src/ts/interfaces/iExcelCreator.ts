import {ExportParams} from "../exporter/exportParams";

export interface ExcelExportParams extends ExportParams<Object[][]> {
    sheetName?: string;
    suppressTextAsCDATA?:boolean;
    exportMode?: "xlsx" | "xml";
}

export interface IExcelCreator {
    exportDataAsExcel(params?: ExcelExportParams): void;
    getDataAsExcelXml(params?: ExcelExportParams): string;
}