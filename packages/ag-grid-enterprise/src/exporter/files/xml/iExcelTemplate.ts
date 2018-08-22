import {
    ExcelStyle,
    ExcelWorksheet,
    ExcelColumn,
    ExcelRow,
    ExcelCell,
    XmlElement
} from 'ag-grid-community';

export interface ExcelTemplate {
    getTemplate(styleProperties?: ExcelStyle | ExcelWorksheet | ExcelColumn | ExcelRow | ExcelCell): XmlElement;
}