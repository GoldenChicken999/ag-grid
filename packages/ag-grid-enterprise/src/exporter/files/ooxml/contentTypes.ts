import {Utils} from 'ag-grid-community';
import {ExcelOOXMLTemplate} from 'ag-grid-community';
import contentType from './contentType';

const contentTypes: ExcelOOXMLTemplate = {
    getTemplate() {

        const children = Utils.map([{
            name: 'Default',
            ContentType: 'application/xml',
            Extension: 'xml'
        },{
            name: 'Default',
            ContentType: "application/vnd.openxmlformats-package.relationships+xml",
            Extension: "rels"
        }, {
            name: 'Override',
            ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            PartName: "/xl/worksheets/sheet1.xml"
        }, {
            name: 'Override',
            ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
            PartName: '/xl/workbook.xml'
        }], contentType.getTemplate);

        return {
            name: "Types",
            properties: {
                rawMap: {
                    xmlns: "http://schemas.openxmlformats.org/package/2006/content-types"
                }
            },
            children
        };
    }
};

export default contentTypes;