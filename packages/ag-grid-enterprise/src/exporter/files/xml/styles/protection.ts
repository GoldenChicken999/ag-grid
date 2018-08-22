import {
    ExcelStyle,
    XmlElement
} from 'ag-grid-community';

import {ExcelTemplate} from '../iExcelTemplate';

const protection: ExcelTemplate = {
    getTemplate(styleProperties: ExcelStyle): XmlElement {
        return {
            name: "Protection",
            properties: {
                prefixedAttributes:[{
                    prefix: "ss:",
                    map: {
                        Protected: styleProperties.protection.protected,
                        HideFormula: styleProperties.protection.hideFormula
                    }
                }]
            }
        };
    }
};

export default protection;