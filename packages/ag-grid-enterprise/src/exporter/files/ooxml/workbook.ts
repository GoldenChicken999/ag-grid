import {ExcelOOXMLTemplate} from 'ag-grid-community';
import sheets from './sheets';

const workbook: ExcelOOXMLTemplate = {
    getTemplate() {

        return {
            name: "workbook",
            properties: {
                prefixedAttributes:[{
                    prefix: "xmlns:",
                    map: {
                        r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    },
                }],
                rawMap: {
                    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                }
            },
            children: [sheets.getTemplate()]
        };
    }
};

export default workbook;