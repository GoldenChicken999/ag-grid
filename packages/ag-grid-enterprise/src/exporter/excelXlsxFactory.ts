import {Autowired, Bean, XmlFactory} from 'ag-grid-community';

import contentTypes from './files/ooxml/contentTypes';
import workbook from './files/ooxml/workbook';
import worksheet from './files/ooxml/worksheet';
import relationships from './files/ooxml/relationships';

import {ExcelStyle, ExcelWorksheet} from 'ag-grid-community';

/**
 * See https://www.ecma-international.org/news/TC45_current_work/OpenXML%20White%20Paper.pdf
 */
@Bean('excelXlsxFactory')
export class ExcelXlsxFactory {

    @Autowired('xmlFactory') private xmlFactory: XmlFactory;

    public createExcel(styles: ExcelStyle[], worksheets: ExcelWorksheet[]): string {
        return this.worksheet(worksheets);
    }

    public contentTypes(): string {
        const header = this.xmlFactory.createHeader();
        const body = this.xmlFactory.createXml(contentTypes.getTemplate());

        return `${header}${body}`;
    }

    public rels(): string {
        const header = this.xmlFactory.createHeader({
            encoding: 'UTF-8',
            standalone: 'yes'
        });
        const rs = relationships.getTemplate([{
            Id: 'rId1',
            Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
            Target: 'xl/workbook.xml'
        }]);
        const body = this.xmlFactory.createXml(rs);

        return `${header}${body}`;
    }

    public workbook(): string {
        const header = this.xmlFactory.createHeader({
            encoding: 'UTF-8',
            standalone: 'yes'
        });
        const body = this.xmlFactory.createXml(workbook.getTemplate());

        return `${header}${body}`;
    }

    public workbookRels(): string {
        const header = this.xmlFactory.createHeader();
        const rs = relationships.getTemplate([{
            Id: 'rId3',
            Target: 'worksheets/sheet1.xml',
            Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
        }]);
        const body = this.xmlFactory.createXml(rs);

        return `${header}${body}`;
    }

    public worksheet(worksheets: ExcelWorksheet[]): string {
        const header = this.xmlFactory.createHeader({
            encoding: 'UTF-8',
            standalone: 'yes'
        });
        const template = this.xmlFactory.createXml(worksheet.getTemplate(worksheets[0]));

        return `${header}${template}`;
    }
}