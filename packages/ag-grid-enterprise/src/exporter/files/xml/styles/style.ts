import {
    ExcelStyle,
    XmlElement
} from 'ag-grid-community';

import {ExcelTemplate} from '../iExcelTemplate';

const style: ExcelTemplate = {
    getTemplate(styleProperties: ExcelStyle): XmlElement {
        const {id, name} = styleProperties;
        return {
            name: 'Style',
            properties: {
                prefixedAttributes:[{
                    prefix: "ss:",
                    map: {
                        ID: id,
                        Name: name ?  name : id
                    }
                }]
            }
        };
    }
};

export default style;