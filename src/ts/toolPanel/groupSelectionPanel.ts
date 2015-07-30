/// <reference path="../widgets/agList.ts" />
/// <reference path="../utils.ts" />
/// <reference path="../svgFactory.ts" />
/// <reference path="../layout/BorderLayout.ts" />
/// <reference path="../constants.ts" />

module awk.grid {

    var _ = Utils;
    var svgFactory = SvgFactory.getInstance();
    var constants = Constants;

    export class GroupSelectionPanel {

        gridOptionsWrapper: any;
        columnController: any;
        inMemoryRowController: any;
        cColumnList: any;
        layout: any;

        constructor(columnController: any, inMemoryRowController: any, gridOptionsWrapper: any) {
            this.gridOptionsWrapper = gridOptionsWrapper;
            this.setupComponents();
            this.columnController = columnController;
            this.inMemoryRowController = inMemoryRowController;

            var that = this;
            this.columnController.addListener({
                columnsChanged: that.columnsChanged.bind(that)
            });
        }

        private columnsChanged(newColumns: any, newGroupedColumns: any) {
            this.cColumnList.setModel(newGroupedColumns);
        }

        public addDragSource(dragSource: any) {
            this.cColumnList.addDragSource(dragSource);
        }

        private columnCellRenderer(params: any) {
            var column = params.value;
            var colDisplayName = this.columnController.getDisplayNameForCol(column);

            var eResult = document.createElement('span');

            var eRemove = _.createIcon('columnRemoveFromGroup',
                this.gridOptionsWrapper, column, svgFactory.createArrowUpSvg);
            _.addCssClass(eRemove, 'ag-visible-icons');
            eResult.appendChild(eRemove);

            var that = this;
            eRemove.addEventListener('click', function () {
                that.columnController.removePivotColumn(column);
            });

            var eValue = document.createElement('span');
            eValue.innerHTML = colDisplayName;
            eResult.appendChild(eValue);

            return eResult;
        }

        private setupComponents() {
            var localeTextFunc = this.gridOptionsWrapper.getLocaleTextFunc();
            var columnsLocalText = localeTextFunc('pivotedColumns', 'Pivoted Columns');
            var pivotedColumnsEmptyMessage = localeTextFunc('pivotedColumnsEmptyMessage', 'Drag columns from above to pivot');

            this.cColumnList = new AgList();
            this.cColumnList.setCellRenderer(this.columnCellRenderer.bind(this));
            //this.cColumnList.addModelChangedListener(this.onGroupingChanged.bind(this));
            this.cColumnList.addBeforeDropListener(this.onBeforeDrop.bind(this));
            this.cColumnList.addItemMovedListener(this.onItemMoved.bind(this));
            this.cColumnList.setEmptyMessage(pivotedColumnsEmptyMessage);
            this.cColumnList.addStyles({height: '100%', overflow: 'auto'});
            this.cColumnList.setReadOnly(true);

            var eNorthPanel = document.createElement('div');
            eNorthPanel.style.paddingTop = '10px';
            eNorthPanel.innerHTML = '<div style="text-align: center;">' + columnsLocalText + '</div>';

            this.layout = new BorderLayout({
                center: this.cColumnList.getGui(),
                north: eNorthPanel
            });
        }

        private onBeforeDrop(newItem: any) {
            this.columnController.addPivotColumn(newItem);
        }

        private onItemMoved(fromIndex: number, toIndex: number) {
            this.columnController.movePivotColumn(fromIndex, toIndex);
        }

        //private onGroupingChanged() {
        //    this.inMemoryRowController.doGrouping();
        //    this.inMemoryRowController.updateModel(constants.STEP_EVERYTHING);
        //    this.columnController.onColumnStateChanged();
        //}
    }
}
