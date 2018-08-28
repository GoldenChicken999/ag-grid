import {Utils} from 'ag-grid-community';
import {ExcelOOXMLTemplate, ExcelWorksheet} from 'ag-grid-community';
import column from './column';
import row from './row';

const worksheet: ExcelOOXMLTemplate = {
    getTemplate(config: ExcelWorksheet) {
        const {table} = config;
        const {rows, columns} = table;

        const children = [].concat(
            columns.length ? {
                name: 'cols',
                children: Utils.map(columns, column.getTemplate)
            } : []
        ).concat(
            rows.length ? {
                name: 'sheetData',
                children: Utils.map(rows, row.getTemplate)
            } : []
        );

        return {
            name: "worksheet",
            properties: {
                prefixedAttributes:[{
                    prefix: "xmlns:",
                    map: {
                        r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    }
                }],
                rawMap: {
                    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                }
            },
            children
        };
    }
};

export default worksheet;