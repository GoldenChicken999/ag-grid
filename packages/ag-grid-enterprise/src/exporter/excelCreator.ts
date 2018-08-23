import {
    Autowired,
    BaseCreator,
    Bean,
    Column,
    ColumnController,
    Downloader,
    ExcelExportParams,
    GridOptions,
    GridOptionsWrapper,
    GridSerializer,
    IExcelCreator,
    PostConstruct,
    RowNode,
    RowType,
    StylingService,
    ValueService,
    _
} from "ag-grid-community";

import {ExcelCell, ExcelStyle} from './interfaces/iExcel';
import {ExcelGridSerializingSession} from './excelGridSerializingSession';
import {ExcelXmlFactory} from "./excelXmlFactory";
import {XlsxFactory} from "./xlsxFactory";
import * as JSZip from 'jszip-sync';

export interface ExcelMixedStyle {
    key: string;
    excelID: string;
    result: ExcelStyle;
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