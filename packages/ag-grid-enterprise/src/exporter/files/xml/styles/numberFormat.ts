import {
    ExcelStyle,
    XmlElement
} from 'ag-grid-community';

import {ExcelTemplate} from '../iExcelTemplate';

const numberFormat: ExcelTemplate = {
    getTemplate(styleProperties: ExcelStyle): XmlElement {
        const {format} = styleProperties.numberFormat;
        return {
            name: "NumberFormat",
            properties: {
                prefixedAttributes:[{
                    prefix: "ss:",
                    map: {
                        Format: format
                    }
                }]
            }
        };
    }
};

export default numberFormat;