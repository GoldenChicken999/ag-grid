import {XmlElement, ExcelWorksheet, Utils} from 'ag-grid-community';
import {ExcelTemplate} from './iExcelTemplate';
import column from './column';
import row from './row';

const worksheet: ExcelTemplate = {
    getTemplate(ws: ExcelWorksheet): XmlElement {
        const {table, name} = ws;
        const {columns, rows} = table;

        const c = Utils.map(columns, (it):XmlElement => column.getTemplate(it));
        const r = Utils.map(rows, (it):XmlElement => row.getTemplate(it));

        return {
            name: "Worksheet",
            children:[{
                name:"Table",
                children: c.concat(r)
            }],
            properties:{
                prefixedAttributes: [{
                    prefix:"ss:",
                    map: {
                        Name: name
                    }
                }]
            }
        };
    }
};

export default worksheet;