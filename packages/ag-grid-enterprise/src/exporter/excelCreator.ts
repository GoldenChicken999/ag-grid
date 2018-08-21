import {
    Autowired,
    BaseCreator,
    BaseGridSerializingSession,
    Bean,
    Column,
    ColumnController,
    Constants,
    Downloader,
    ExcelCell,
    ExcelColumn,
    ExcelDataType,
    ExcelExportParams,
    ExcelRow,
    ExcelStyle,
    ExcelWorksheet,
    GridOptions,
    GridOptionsWrapper,
    GridSerializer,
    IExcelCreator,
    PostConstruct,
    ProcessCellForExportParams,
    ProcessHeaderForExportParams,
    RowAccumulator,
    RowNode,
    RowSpanningAccumulator,
    RowType,
    StylingService,
    Utils,
    ValueService,
    _
} from "ag-grid-community";

import {ExcelXmlFactory} from "./excelXmlFactory";
import {XlsxFactory} from "./xlsxFactory";
import * as JSZip from 'jszip-sync';

export interface ExcelMixedStyle {
    key: string;
    excelID: string;
    result: ExcelStyle;
}

export class ExcelGridSerializingSession extends BaseGridSerializingSession<ExcelCell[][]> {
    private stylesByIds: any;
    private mixedStyles: { [key: string]: ExcelMixedStyle } = {};
    private mixedStyleCounter: number = 0;
    private excelStyles: ExcelStyle[];
    private customHeader: ExcelCell[][];
    private customFooter: ExcelCell[][];
    private sheetName:string;
    private suppressTextAsCDATA:boolean;

    constructor(columnController: ColumnController,
                valueService: ValueService,
                gridOptionsWrapper: GridOptionsWrapper,
                processCellCallback: (params: ProcessCellForExportParams) => string,
                processHeaderCallback: (params: ProcessHeaderForExportParams) => string,
                sheetName:string,
                private excelFactory: XlsxFactory | ExcelXmlFactory,
                baseExcelStyles: ExcelStyle[],
                private styleLinker: (rowType: RowType, rowIndex: number, colIndex: number, value: string, column: Column, node: RowNode) => string[],
                suppressTextAsCDATA:boolean) {
        super(columnController, valueService, gridOptionsWrapper, processCellCallback, processHeaderCallback, (raw: string) => raw);
        this.stylesByIds = {};

        if (!baseExcelStyles) {
            this.excelStyles = [];
        } else {
            baseExcelStyles.forEach((it: ExcelStyle) => {
                this.stylesByIds[it.id] = it;
            });
            this.excelStyles = baseExcelStyles.slice();
        }
        this.sheetName = sheetName;
        this.suppressTextAsCDATA = suppressTextAsCDATA;
    }

    private rows: ExcelRow[] = [];

    private cols: ExcelColumn[];

    public addCustomHeader(customHeader: ExcelCell[][]): void {
        this.customHeader = customHeader;
    }

    public addCustomFooter(customFooter: ExcelCell[][]): void {
        this.customFooter = customFooter;
    }

    public prepare(columnsToExport: Column[]): void {
        this.cols = Utils.map(columnsToExport, (it: Column) => {
            it.getColDef().cellStyle;
            return {
                width: it.getActualWidth()
            };
        });
    }

    public onNewHeaderGroupingRow(): RowSpanningAccumulator {
        let currentCells: ExcelCell[] = [];
        let that = this;
        this.rows.push({
            cells: currentCells
        });
        return {
            onColumn: (header: string, index: number, span: number) => {
                let styleIds: string[] = that.styleLinker(RowType.HEADER_GROUPING, 1, index, "grouping-" + header, null, null);
                currentCells.push(that.createMergedCell(styleIds.length > 0 ? styleIds[0] : null, "String", header, span));
            }
        };
    }

    public onNewHeaderRow(): RowAccumulator {
        return this.onNewRow(this.onNewHeaderColumn);
    }

    public onNewBodyRow(): RowAccumulator {
        return this.onNewRow(this.onNewBodyColumn);
    }

    onNewRow(onNewColumnAccumulator: (rowIndex: number, currentCells: ExcelCell[]) => (column: Column, index: number, node?: RowNode) => void): RowAccumulator {
        let currentCells: ExcelCell[] = [];
        this.rows.push({
            cells: currentCells
        });
        return {
            onColumn: onNewColumnAccumulator.bind(this, this.rows.length, currentCells)()
        };
    }

    onNewHeaderColumn(rowIndex: number, currentCells: ExcelCell[]): (column: Column, index: number, node?: RowNode) => void {
        let that = this;
        return (column: Column, index: number, node?: RowNode) => {
            let nameForCol = this.extractHeaderValue(column);
            let styleIds: string[] = that.styleLinker(RowType.HEADER, rowIndex, index, nameForCol, column, null);
            currentCells.push(this.createCell(styleIds.length > 0 ? styleIds[0] : null, 'String', nameForCol));
        };
    }

    public parse(): string {
        function join(header: ExcelCell[][], body: ExcelRow[], footer: ExcelCell[][]): ExcelRow[] {
            let all: ExcelRow[] = [];
            if (header) {
                header.forEach(rowArray => all.push({cells: rowArray}));
            }
            body.forEach(it => all.push(it));
            if (footer) {
                footer.forEach(rowArray => all.push({cells: rowArray}));
            }
            return all;
        }

        let data: ExcelWorksheet [] = [{
            name: this.sheetName,
            table: {
                columns: this.cols,
                rows: join(this.customHeader, this.rows, this.customFooter)
            }
        }];

        return this.excelFactory.createExcel(this.excelStyles, data);
    }

    onNewBodyColumn(rowIndex: number, currentCells: ExcelCell[]): (column: Column, index: number, node?: RowNode) => void {
        let that = this;
        return (column: Column, index: number, node?: RowNode) => {
            let valueForCell = this.extractRowCellValue(column, index, Constants.EXPORT_TYPE_EXCEL, node);
            let styleIds: string[] = that.styleLinker(RowType.BODY, rowIndex, index, valueForCell, column, node);
            let excelStyleId: string = null;
            if (styleIds && styleIds.length == 1) {
                excelStyleId = styleIds [0];
            } else if (styleIds && styleIds.length > 1) {
                let key: string = styleIds.join("-");
                if (!this.mixedStyles[key]) {
                    this.addNewMixedStyle(styleIds);
                }
                excelStyleId = this.mixedStyles[key].excelID;
            }
            let type: ExcelDataType = Utils.isNumeric(valueForCell) ? 'Number' : 'String';
            currentCells.push(that.createCell(excelStyleId, type, valueForCell));
        };
    }

    addNewMixedStyle(styleIds: string[]): void {
        this.mixedStyleCounter += 1;
        let excelId = 'mixedStyle' + this.mixedStyleCounter;
        let resultantStyle: ExcelStyle = {};

        styleIds.forEach((styleId: string) => {
            this.excelStyles.forEach((excelStyle: ExcelStyle) => {
                if (excelStyle.id === styleId) {
                    Utils.mergeDeep(resultantStyle, excelStyle);
                }
            });
        });

        resultantStyle['id'] = excelId;
        resultantStyle['name'] = excelId;
        let key: string = styleIds.join("-");
        this.mixedStyles[key] = {
            excelID: excelId,
            key: key,
            result: resultantStyle
        };
        this.excelStyles.push(resultantStyle);
        this.stylesByIds[excelId] = resultantStyle;
    }

    private styleExists(styleId: string): boolean {
        if (styleId == null) return false;

        return this.stylesByIds[styleId];
    }

    private createCell(styleId: string, type: ExcelDataType, value: string): ExcelCell {
        let actualStyle: ExcelStyle = this.stylesByIds[styleId];
        let styleExists: boolean = actualStyle != null;

        function getType(): ExcelDataType {
            if (
                styleExists &&
                actualStyle.dataType
            ) switch (actualStyle.dataType) {
                case 'string':
                    return 'String';
                case 'number':
                    return 'Number';
                case 'dateTime':
                    return 'DateTime';
                case 'error':
                    return 'Error';
                case 'boolean':
                    return 'Boolean';
                default:
                    console.warn(`ag-grid: Unrecognized data type for excel export [${actualStyle.id}.dataType=${actualStyle.dataType}]`);
            }

            return type;
        }

        let typeTransformed: ExcelDataType = getType();

        let massageText = (val:string) =>  this.suppressTextAsCDATA ? _.escape(val) : `<![CDATA[${val}]]>`;

        return {
            styleId: styleExists ? styleId : null,
            data: {
                type: typeTransformed,
                value:
                    typeTransformed === 'String' ? massageText(value):
                    typeTransformed === 'Number' ? Number(value).valueOf() + '' :
                    value
            }
        };
    }

    private createMergedCell(styleId: string, type: ExcelDataType, value: string, numOfCells: number): ExcelCell {
        return {
            styleId: this.styleExists(styleId) ? styleId : null,
            data: {
                type: type,
                value: value
            },
            mergeAcross: numOfCells
        };
    }
}

@Bean('excelCreator')
export class ExcelCreator extends BaseCreator<ExcelCell[][], ExcelGridSerializingSession, ExcelExportParams> implements IExcelCreator {

    @Autowired('excelXmlFactory') private excelXmlFactory: ExcelXmlFactory;
    @Autowired('xlsxFactory') private xlsxFactory: XlsxFactory;
    @Autowired('columnController') private columnController: ColumnController;
    @Autowired('valueService') private valueService: ValueService;
    @Autowired('gridOptions') private gridOptions: GridOptions;
    @Autowired('stylingService') private stylingService: StylingService;

    @Autowired('downloader') private downloader: Downloader;
    @Autowired('gridSerializer') private gridSerializer: GridSerializer;
    @Autowired('gridOptionsWrapper') gridOptionsWrapper: GridOptionsWrapper;

    private exportMode: string;

    @PostConstruct
    public postConstruct(): void {
        this.setBeans({
            downloader: this.downloader,
            gridSerializer: this.gridSerializer,
            gridOptionsWrapper: this.gridOptionsWrapper
        });
    }

    public exportDataAsExcel(params?: ExcelExportParams): string {
        if (params.exportMode) {
            this.setExportMode(params.exportMode);
        }
        return this.export(params);
    }

    public getDataAsExcelXml(params?: ExcelExportParams): string {
        return this.getData(params);
    }

    public getMimeType(): string {
        return this.getExportMode() === 'xml' ? 'application/vnd.ms-excel' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    }

    public getDefaultFileName(): string {
        return `export.${this.getExportMode()}`;
    }

    public getDefaultFileExtension(): string {
        return this.getExportMode();
    }

    public createSerializingSession(params?: ExcelExportParams): ExcelGridSerializingSession {
        const factory = this.getExportMode() === 'xml' ? this.excelXmlFactory : this.xlsxFactory;
        return new ExcelGridSerializingSession(
            this.columnController,
            this.valueService,
            this.gridOptionsWrapper,
            params ? params.processCellCallback : null,
            params ? params.processHeaderCallback : null,
            params && params.sheetName != null && params.sheetName != "" ? params.sheetName : 'ag-grid',
            factory,
            this.gridOptions.excelStyles,
            this.styleLinker.bind(this),
            params && params.suppressTextAsCDATA ? params.suppressTextAsCDATA : false
        );
    }

    private styleLinker(rowType: RowType, rowIndex: number, colIndex: number, value: string, column: Column, node: RowNode): string[] {
        if ((rowType === RowType.HEADER) || (rowType === RowType.HEADER_GROUPING)) return ["header"];
        if (!this.gridOptions.excelStyles || this.gridOptions.excelStyles.length === 0) return null;

        let styleIds: string[] = this.gridOptions.excelStyles.map((it: ExcelStyle) => {
            return it.id;
        });

        let applicableStyles: string [] = [];
        this.stylingService.processAllCellClasses(
            column.getColDef(),
            {
                value: value,
                data: node.data,
                node: node,
                colDef: column.getColDef(),
                rowIndex: rowIndex,
                api: this.gridOptionsWrapper.getApi(),
                context: this.gridOptionsWrapper.getContext()
            },
            (className: string) => {
                if (styleIds.indexOf(className) > -1) {
                    applicableStyles.push(className);
                }
            }
        );

        return applicableStyles.sort((left: string, right: string): number => {
            return (styleIds.indexOf(left) < styleIds.indexOf(right)) ? -1 : 1;
        });
    }

    public isExportSuppressed():boolean {
        return this.gridOptionsWrapper.isSuppressExcelExport();
    }

    private setExportMode(exportMode: string): void {
        this.exportMode = exportMode;
    }

    private getExportMode(): string {
        return this.exportMode || 'xlsx';
    }

    protected packageFile(data: string): Blob {
        if (this.getExportMode() === 'xml') {
            return super.packageFile(data);
        }

        const zip: JSZip = new JSZip();
        const xlsxFactory = this.xlsxFactory;

        return zip.sync(() => {
            const xl = zip.folder('xl');
            xl.file('workbook.xml', xlsxFactory.workbook());
            xl.file('_rels/workbook.xml.rels', xlsxFactory.workbookRels());
            zip.file('_rels/.rels', xlsxFactory.rels());
            zip.file('[Content_Types].xml', xlsxFactory.contentTypes());

            xl.file('worksheets/sheet1.xml', data);

            let zipped;

            zip.generateAsync({
                type: 'blob',
                mimeType:
                  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
              }).then(function(content: any) {
                zipped = content;
            });

            return zipped;
        });

    }

}