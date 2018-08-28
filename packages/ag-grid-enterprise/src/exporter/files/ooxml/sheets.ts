import {ExcelOOXMLTemplate} from 'ag-grid-community';
import sheet from './sheet';

const sheets: ExcelOOXMLTemplate = {
    getTemplate() {
        return {
            name: "sheets",
            children: [sheet.getTemplate()]
        };
    }
};

export default sheets;