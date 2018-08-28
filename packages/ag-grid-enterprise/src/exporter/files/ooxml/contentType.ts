import {ExcelContentType, ExcelOOXMLTemplate} from 'ag-grid-community';

const contentType: ExcelOOXMLTemplate = {
    getTemplate(config: ExcelContentType) {
        const {name, ContentType, Extension, PartName} = config;

        return {
            name,
            properties: {
                rawMap: {
                    ContentType,
                    Extension,
                    PartName
                }
            }
        };
    }
};

export default contentType;