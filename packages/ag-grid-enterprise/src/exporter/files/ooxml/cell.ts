import {Utils} from 'ag-grid-community';
import {ExcelOOXMLTemplate, ExcelCell} from '../../interfaces/iExcel';

const cell: ExcelOOXMLTemplate = {
    getTemplate(config: ExcelCell) {
        const {ref, data, styleId} = config;
        const {type, value} = data;
        const obj = {
            name: 'c',
            properties: {
                rawMap: {
                    r: ref,
                    t: type,
                    s: styleId
                }
            }
        };

        let children;

        if (type === 'inlineStr') {
            children = [{
                name: 'is',
                children: [{
                    name: 't',
                    textNode: value
                }]
            }]
        } else {
            children = [{
                name: 'v',
                textNode: value
            }]
        }

        return Utils.assign({}, obj, { children });
    }
};

export default cell;