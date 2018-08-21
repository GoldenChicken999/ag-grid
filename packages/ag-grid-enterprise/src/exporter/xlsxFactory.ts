import {Autowired, Bean, ExcelStyle, ExcelWorksheet, XmlElement, XmlFactory} from 'ag-grid-community';
import {ExcelXmlFactory} from './excelXmlFactory';

/**
 * See https://www.ecma-international.org/news/TC45_current_work/OpenXML%20White%20Paper.pdf
 */
@Bean('xlsxFactory')
export class XlsxFactory {

    @Autowired('xmlFactory') private xmlFactory: XmlFactory;
    @Autowired('excelXmlFactory') private excelXmlFactory: ExcelXmlFactory;

    public createExcel(styles: ExcelStyle[], worksheets: ExcelWorksheet[]): string {
        return this.worksheet(styles, worksheets);
    }

    public contentTypes(): string {
        const header = '<?xml version="1.0" ?>';
        const body = this.xmlFactory.createXml({
            name: "Types",
            properties: {
                rawMap: {
                    xmlns: "http://schemas.openxmlformats.org/package/2006/content-types"
                }
            },
            children: [{
                name: "Default",
                properties: {
                    rawMap: {
                        ContentType: "application/xml",
                        Extension: "xml"
                    }
                }
            }, {
                name: "Default",
                properties: {
                    rawMap: {
                        ContentType: "application/vnd.openxmlformats-package.relationships+xml",
                        Extension: "rels"
                    }
                }
            }, {
                name: "Override",
                properties: {
                    rawMap: {
                        ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                        PartName: "/xl/worksheets/sheet1.xml"
                    }
                }
            }, {
                name: "Override",
                properties: {
                    rawMap: {
                        ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                        PartName: "/xl/workbook.xml"
                    }
                }
            }]
        });

        return `${header}${body}`;
    }

    public rels(): string {
        const header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        const body = this.xmlFactory.createXml({
            name: "Relationships",
            properties: {
                rawMap: {
                    xmlns: "http://schemas.openxmlformats.org/package/2006/relationships"
                }
            },
            children: [{
                name: "Relationship",
                properties: {
                    rawMap: {
                        Id: "rId1",
                        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                        Target: "xl/workbook.xml"
                    }
                }
            }]
        });

        return `${header}${body}`;
    }

    public workbook(): string {
        const header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        const body = this.xmlFactory.createXml({
            name: "workbook",
            properties: {
                prefixedAttributes:[{
                    prefix: "xmlns:",
                    map: {
                        r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                        mx: "http://schemas.microsoft.com/office/mac/excel/2008/main",
                        mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
                        mv: "urn:schemas-microsoft-com:mac:vml",
                        x14: "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
                        x14ac: "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
                        xm: "http://schemas.microsoft.com/office/excel/2006/main"
                    },
                }],
                rawMap: {
                    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                }
            },
            children: [{
                name: "workbookPr"
            }, {
                name: "sheets",
                children: [{
                    name: "sheet",
                    properties: {
                        rawMap: {
                            "state": "visible",
                            "name": "Sheet1",
                            "sheetId": "1",
                            "r:id": "rId3"
                        }
                    }
                }]
            }, {
                name: "definedNames"
            }, {
                name: "calcPr"
            }]
        });

        return `${header}${body}`;
    }

    public workbookRels(): string {
        const header = '<?xml version="1.0" ?>';
        const body = this.xmlFactory.createXml({
            name: "Relationships",
            properties: {
                rawMap: {
                    xmlns: "http://schemas.openxmlformats.org/package/2006/relationships"
                }
            },
            children: [{
                name: "Relationship",
                properties: {
                    rawMap: {
                        Id: "rId3",
                        Target: "worksheets/sheet1.xml",
                        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                    }
                }
            }]
        });

        return `${header}${body}`;
    }

    public worksheet(styles: ExcelStyle[], worksheets: ExcelWorksheet[]): string {
        const header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        const template = this.xmlFactory.createXml({
            name: "worksheet",
            properties: {
                prefixedAttributes:[{
                    prefix: "xmlns:",
                    map: {
                        mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
                        mv: "urn:schemas-microsoft-com:mac:vml",
                        mx: "http://schemas.microsoft.com/office/mac/excel/2008/main",
                        r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                        x14: "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
                        x14ac: "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
                        xm: "http://schemas.microsoft.com/office/excel/2006/main"
                    }
                }],
                rawMap: {
                    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                }
            },
            children: [{
                name: "sheetData",
                children: [{
                    name: "row",
                    properties: {
                        rawMap: {
                            r: 1
                        }
                    },
                    children: [{
                        name: "c",
                        properties: {
                            rawMap: {
                                r:"A1",
                                t: "inlineStr"
                            }
                        },
                        children: [{
                            name: "is",
                            children: [{
                                name: "t",
                                textNode: "Foo"
                            }]
                        }]
                    }, {
                        name: "c",
                        properties: {
                            rawMap: {
                                r: "B1"
                            }
                        },
                        children: [{
                            name: "v",
                            textNode: "1000"
                        }]
                    }]
                }]
            }]
        });
        return `${header}${template}`;
    }
}