import {ExcelCell, ExcelRow, XmlElement, Utils} from 'ag-grid-community';
import {ExcelTemplate} from './iExcelTemplate';
import cell from './cell';

const row: ExcelTemplate = {
    getTemplate(r: ExcelRow): XmlElement {
        const {cells} = r;

        return {
            name: "Row",
            children: Utils.map(cells, (it:ExcelCell):XmlElement => {
                return cell.getTemplate(it);
            })
        };
    }
};

export default row;