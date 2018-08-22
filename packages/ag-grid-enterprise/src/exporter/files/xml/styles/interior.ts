import {
    ExcelStyle,
    XmlElement
} from 'ag-grid-community';

import {ExcelTemplate} from '../iExcelTemplate';

const interior: ExcelTemplate = {
    getTemplate(styleProperties: ExcelStyle): XmlElement {
        const {color, pattern, patternColor} = styleProperties.interior;
        return {
            name: "Interior",
            properties: {
                prefixedAttributes:[{
                    prefix: "ss:",
                    map: {
                        Color: color,
                        Pattern: pattern,
                        PatternColor: patternColor
                    }
                }]
            }
        };
    }
};

export default interior;