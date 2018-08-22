import {
    ExcelStyle,
    XmlElement
} from 'ag-grid-community';

import {ExcelTemplate} from '../iExcelTemplate';

const alignment: ExcelTemplate = {
    getTemplate(styleProperties: ExcelStyle): XmlElement {
        const {
            vertical,
            horizontal,
            indent,
            readingOrder,
            rotate,
            shrinkToFit,
            verticalText,
            wrapText
        } = styleProperties.alignment;
        return {
            name: 'Alignment',
            properties: {
                prefixedAttributes:[{
                    prefix: "ss:",
                    map: {
                        Vertical: vertical,
                        Horizontal: horizontal,
                        Indent: indent,
                        ReadingOrder: readingOrder,
                        Rotate: rotate,
                        ShrinkToFit: shrinkToFit,
                        VerticalText:verticalText,
                        WrapText: wrapText
                    }
                }]
            }
        };
    }
};

export default alignment;