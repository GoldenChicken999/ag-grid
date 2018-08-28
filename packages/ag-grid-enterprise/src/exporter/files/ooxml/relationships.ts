import {Utils} from 'ag-grid-community';
import {ExcelOOXMLTemplate, ExcelRelationship} from 'ag-grid-community';
import relationship from './relationship';

const relationships: ExcelOOXMLTemplate = {
    getTemplate(c: ExcelRelationship[]) {
        const children = Utils.map(c, relationship.getTemplate);

        return {
            name: "Relationships",
            properties: {
                rawMap: {
                    xmlns: "http://schemas.openxmlformats.org/package/2006/relationships"
                }
            },
            children
        };
    }
};

export default relationships;