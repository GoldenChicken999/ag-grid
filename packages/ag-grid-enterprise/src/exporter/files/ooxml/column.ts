import {ExcelOOXMLTemplate, ExcelColumn} from 'ag-grid-community';

const column: ExcelOOXMLTemplate = {
    getTemplate(config: ExcelColumn) {
        const {min, max, width = 10, s, hidden, bestFit} = config;

        return {
            name: 'col',
            properties: {
                rawMap: {
                    min,
                    max,
                    width,
                    style: s,
                    hidden: hidden ? '1' : '0',
                    bestFit: bestFit ? '1' : '0',
                    customWidth: width !== 10 ? '1' : '0'
                }
            }
        };
    }
};

export default column;