import {XmlElement, Utils} from 'ag-grid-community';
import {ExcelCell, ExcelRow, ExcelXMLTemplate} from '../../interfaces/iExcel';
import cell from './cell';

const row: ExcelXMLTemplate = {
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