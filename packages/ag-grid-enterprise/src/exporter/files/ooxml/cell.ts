import {Utils} from 'ag-grid-community';
import {ExcelOOXMLTemplate, ExcelCell} from 'ag-grid-community';

const cell: ExcelOOXMLTemplate = {
    getTemplate(config: ExcelCell) {
        const {ref, data, styleId} = config;
        const {type, value} = data;
        const processedType = cell.convertType(type);
        const obj = {
            name: 'c',
            properties: {
                rawMap: {
                    r: ref,
                    t: processedType,
                    //s: styleId
                }
            }
        };

        let children;

        if (processedType === 'inlineStr') {
            children = [{
                name: 'is',
                children: [{
                    name: 't',
                    textNode: value
                }]
            }];
        } else {
            children = [{
                name: 'v',
                textNode: value
            }];
        }

        return Utils.assign({}, obj, { children });
    },

    convertType(type) {
        let t;

        switch(type) {
            case 'String':
                t = 'inlineStr';
                break;
            case 'Number':
                t = 'n';
                break;
            case 'Boolean':
                t = 'b';
                break;
            case 'DateTime':
                t = 'd';
                break;
            case 'Error':
                t = 'e';
            default:
                t = type;
        }

        return t;
    }
};

export default cell;