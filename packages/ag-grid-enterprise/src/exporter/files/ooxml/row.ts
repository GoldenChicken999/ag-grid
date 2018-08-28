import {Utils} from 'ag-grid-community';
import {ExcelOOXMLTemplate, ExcelRow} from 'ag-grid-community';
import cell from './cell';

const row: ExcelOOXMLTemplate = {
    getTemplate(config: ExcelRow) {
        const {index, collapsed, hidden, height, outlineLevel, s, cells = []} = config;
        const children = Utils.map(cells, cell.getTemplate);
        return {
            name: "row",
            properties: {
                rawMap: {
                    r: index,
                    collapsed,
                    hidden: hidden ? '1' : '0',
                    ht: height,
                    customHeight: height != null ? '1' : '0',
                    s,
                    customFormat: s != null ? '1' : '0'
                }
            },
            children
        };
    }
};

export default row;