import {XmlElement, ExcelColumn} from 'ag-grid-community';
import {ExcelTemplate} from './iExcelTemplate';

const column: ExcelTemplate = {
    getTemplate(c: ExcelColumn): XmlElement {
        const {width} = c;
        return {
            name:"Column",
            properties:{
                prefixedAttributes: [{
                    prefix:"ss:",
                    map: {
                        Width: width
                    }
                }]
            }
        };
    }
};

export default column;