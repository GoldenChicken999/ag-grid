import {ExcelContentType, ExcelOOXMLTemplate} from '../../interfaces/iExcel';

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