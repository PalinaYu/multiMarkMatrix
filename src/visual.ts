"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";
import "./../style/visual.less";

import VisualConstructorOptions =
    powerbi.extensibility.visual.VisualConstructorOptions;

import VisualUpdateOptions =
    powerbi.extensibility.visual.VisualUpdateOptions;

import IVisual =
    powerbi.extensibility.visual.IVisual;

import IVisualHost =
    powerbi.extensibility.visual.IVisualHost;

import IVisualEventService =
    powerbi.extensibility.IVisualEventService;

import ISelectionManager =
    powerbi.extensibility.ISelectionManager;

import ISelectionId =
    powerbi.visuals.ISelectionId;

import ILocalizationManager =
    powerbi.extensibility.ILocalizationManager;

import VisualTooltipDataItem =
    powerbi.extensibility.VisualTooltipDataItem;

import { VisualFormattingSettingsModel } from "./settings";


interface MatrixAxisItem {
    key: string;
    values: string[];
    sortValues?: unknown[];
}


interface MarkPoint {
    status: string;
    order: string;
    selectionId: ISelectionId;
    tooltipItems: VisualTooltipDataItem[];
}


interface PersistedLayoutState {
    version: number;
    rowWidths: Array<[number, number]>;
    dataWidths: Array<[string, number]>;
}


/*
 * Power BI currently exposes rolesIndex at runtime for table mappings,
 * even though it is not part of the public DataViewMetadataColumn type.
 *
 * We keep it isolated behind this narrow optional interface and always
 * fall back to metadata order if the host does not provide it.
 */
interface RuntimeRoleIndexColumn
    extends powerbi.DataViewMetadataColumn {

    rolesIndex?: {
        [roleName: string]:
            number[] | undefined;
    };
}


export class Visual implements IVisual {

    private host: IVisualHost;
    private events: IVisualEventService;
    private selectionManager: ISelectionManager;
    private target: HTMLElement;
    private locale: string;
    private localizationManager: ILocalizationManager;

    /*
     * Power BI tells the visual whether user interaction is allowed
     * in the current host scenario.
     */
    private allowInteractions:
        boolean =
        true;

    /*
     * Power BI / Windows high-contrast palette.
     * These values come from the host, not from hard-coded colors.
     */
    private isHighContrast:
        boolean =
        false;

    private highContrastForeground:
        string =
        "#000000";

    private highContrastBackground:
        string =
        "#FFFFFF";

    private highContrastSelected:
        string =
        "#000000";

    private formattingSettingsService: FormattingSettingsService;

    private formattingSettings:
        VisualFormattingSettingsModel =
        new VisualFormattingSettingsModel();


    /*
     * Manual column widths.
     * They survive normal visual updates while this visual instance is alive.
     */
    private rowColumnWidths: number[] = [];

    private dataColumnWidths:
        Map<string, number> =
        new Map<string, number>();


    /*
     * Auto width is recalculated from content on every update
     * until the user manually resizes a column.
     */
    private manuallyResizedRowColumns:
        Set<number> =
        new Set<number>();

    private manuallyResizedDataColumns:
        Set<string> =
        new Set<string>();


    /*
     * State used to synchronize persisted column widths with Power BI.
     * Widths themselves are stored in internalState.layoutState.
     */
    private lastLayoutStateFromDataView:
        string | null =
        null;

    private pendingLayoutState:
        string | null =
        null;


    /*
     * Per-column-field sort state.
     * If a field is absent from this map it uses Power BI data order.
     */
    private columnFieldSortState:
        Map<string, "ascending" | "descending"> =
        new Map<string, "ascending" | "descending">();

    private lastColumnFieldSortStateFromDataView:
        string | null =
        null;

    private pendingColumnFieldSortState:
        string | null =
        null;


    /*
     * Per-row-field sort state.
     *
     * Modes:
     * - absent  -> Power BI data order
     * - ascending
     * - descending
     * - manual
     *
     * Manual order is stored separately for every parent group,
     * so a child field can be reordered inside one parent without
     * mixing values from another parent.
     */
    private rowFieldSortState:
        Map<
            string,
            "ascending" | "descending" | "manual"
        > =
        new Map<
            string,
            "ascending" | "descending" | "manual"
        >();

    private manualRowOrder:
        Map<
            string,
            Map<string, string[]>
        > =
        new Map<
            string,
            Map<string, string[]>
        >();

    private lastRowSortStateFromDataView:
        string | null =
        null;

    private pendingRowSortState:
        string | null =
        null;


    constructor(options: VisualConstructorOptions) {

        this.host = options.host;
        this.events = options.host.eventService;

        this.selectionManager =
            options.host.createSelectionManager();

        this.target =
            options.element;

        this.locale =
            options.host.locale;


        this.localizationManager =
            options.host
                .createLocalizationManager();


        this.allowInteractions =
            options.host
                .hostCapabilities
                .allowInteractions;


        const colorPalette =
            options.host.colorPalette;


        this.isHighContrast =
            colorPalette.isHighContrast;


        if (
            this.isHighContrast
        ) {

            this.highContrastForeground =
                colorPalette.foreground.value;


            this.highContrastBackground =
                colorPalette.background.value;


            this.highContrastSelected =
                colorPalette.foregroundSelected.value;
        }


        this.formattingSettingsService =
            new FormattingSettingsService(
                this.localizationManager
            );

        this.target.replaceChildren();
    }


    public update(options: VisualUpdateOptions): void {

        this.events.renderingStarted(options);

        try {

            this.target.replaceChildren();

            const dataView =
                options.dataViews?.[0];


            if (!dataView) {

                this.showMessage(
                    "Add fields to Rows, Columns and Category to build the matrix."
                );

                this.events.renderingFinished(
                    options
                );

                return;
            }


            this.formattingSettings =
                this.formattingSettingsService
                    .populateFormattingSettingsModel(
                        VisualFormattingSettingsModel,
                        dataView
                    );


            this.restoreLayoutState(
                dataView
            );


            this.restoreColumnFieldSortState(
                dataView
            );


            this.restoreRowSortState(
                dataView
            );


            const table =
                dataView.table;


            if (
                !table
            ) {

                this.showMessage(
                    "Add fields to Rows, Columns and Category to build the matrix."
                );

                this.events.renderingFinished(
                    options
                );

                return;
            }


            const metadataColumns =
                table.columns;


            const dataRows =
                table.rows ?? [];


            const rowIndexes =
                this.getRoleIndexes(
                    metadataColumns,
                    "rows"
                );


            const columnIndexes =
                this.getRoleIndexes(
                    metadataColumns,
                    "columns"
                );


            const statusIndex =
                this.getRoleIndex(
                    metadataColumns,
                    "status"
                );


            const markOrderIndex =
                this.getRoleIndex(
                    metadataColumns,
                    "markOrder"
                );


            const tooltipIndexes =
                this.getRoleIndexes(
                    metadataColumns,
                    "tooltips"
                );


            if (
                rowIndexes.length === 0 ||
                columnIndexes.length === 0 ||
                statusIndex < 0
            ) {

                this.showMessage(
                    "Add at least one field to Rows, one to Columns, and one to Category."
                );

                this.events.renderingFinished(
                    options
                );

                return;
            }


            const matrixRows =
                new Map<string, MatrixAxisItem>();


            const matrixColumns =
                new Map<string, MatrixAxisItem>();


            const cells =
                new Map<string, MarkPoint[]>();


            dataRows.forEach(
                (
                    dataRow,
                    dataRowIndex
                ) => {


                    const rowValues =
                        rowIndexes.map(
                            index =>
                                this.formatValue(
                                    dataRow[index],
                                    metadataColumns[index]
                                )
                        );


                    const rowSortValues =
                        rowIndexes.map(
                            index =>
                                this.getAxisSortValue(
                                    dataRow[index],
                                    metadataColumns[index]
                                )
                        );


                    const columnValues =
                        columnIndexes.map(
                            index =>
                                this.formatValue(
                                    dataRow[index],
                                    metadataColumns[index]
                                )
                        );


                    const columnSortValues =
                        columnIndexes.map(
                            index =>
                                this.getAxisSortValue(
                                    dataRow[index],
                                    metadataColumns[index]
                                )
                        );


                    const status =
                        this.rawValueToText(
                            dataRow[statusIndex]
                        );


                    const markOrder =
                        markOrderIndex >= 0
                            ? this.rawValueToText(
                                dataRow[
                                    markOrderIndex
                                ]
                            )
                            : String(
                                dataRowIndex
                            );


                    const rowKey =
                        JSON.stringify(
                            rowValues
                        );


                    const columnKey =
                        JSON.stringify(
                            columnValues
                        );


                    const cellKey =
                        rowKey +
                        "___CELL___" +
                        columnKey;


                    if (
                        !matrixRows.has(
                            rowKey
                        )
                    ) {

                        matrixRows.set(
                            rowKey,
                            {
                                key: rowKey,
                                values: rowValues,
                                sortValues:
                                    rowSortValues
                            }
                        );
                    }


                    if (
                        !matrixColumns.has(
                            columnKey
                        )
                    ) {

                        matrixColumns.set(
                            columnKey,
                            {
                                key: columnKey,
                                values: columnValues,
                                sortValues:
                                    columnSortValues
                            }
                        );
                    }


                    const selectionId =
                        this.host
                            .createSelectionIdBuilder()
                            .withTable(
                                table,
                                dataRowIndex
                            )
                            .createSelectionId();


                    const tooltipItems:
                        VisualTooltipDataItem[] =
                        [];


                    tooltipItems.push({

                        displayName:
                            metadataColumns[
                                statusIndex
                            ].displayName,

                        value:
                            status,

                        color:
                            this.getStatusColor(
                                status
                            )
                    });


                    if (
                        markOrderIndex >= 0
                    ) {

                        tooltipItems.push({

                            displayName:
                                metadataColumns[
                                    markOrderIndex
                                ].displayName,

                            value:
                                this.formatValue(
                                    dataRow[
                                        markOrderIndex
                                    ],
                                    metadataColumns[
                                        markOrderIndex
                                    ]
                                )
                        });
                    }


                    tooltipIndexes.forEach(
                        index => {

                            tooltipItems.push({

                                displayName:
                                    metadataColumns[
                                        index
                                    ].displayName,

                                value:
                                    this.formatValue(
                                        dataRow[index],
                                        metadataColumns[index]
                                    )
                            });
                        }
                    );


                    const mark: MarkPoint = {

                        status:
                            status,

                        order:
                            markOrder,

                        selectionId:
                            selectionId,

                        tooltipItems:
                            tooltipItems
                    };


                    if (
                        !cells.has(
                            cellKey
                        )
                    ) {

                        cells.set(
                            cellKey,
                            []
                        );
                    }


                    cells
                        .get(cellKey)!
                        .push(mark);
                }
            );


            cells.forEach(
                marks => {

                    marks.sort(
                        (
                            a,
                            b
                        ) =>
                            this.compareOrder(
                                a.order,
                                b.order
                            )
                    );
                }
            );


            let rows =
                Array.from(
                    matrixRows.values()
                );


            const rowFieldIds =
                rowIndexes.map(
                    (
                        index,
                        level
                    ) =>
                        metadataColumns[
                            index
                        ].queryName ??
                        `rows:${level}:${
                            metadataColumns[
                                index
                            ].displayName
                        }`
                );


            rows =
                this.sortRowsByFieldState(
                    rows,
                    rowFieldIds
                );


            let columns =
                Array.from(
                    matrixColumns.values()
                );


            const columnFieldIds =
                columnIndexes.map(
                    (
                        index,
                        level
                    ) =>
                        metadataColumns[
                            index
                        ].queryName ??
                        `columns:${level}:${
                            metadataColumns[
                                index
                            ].displayName
                        }`
                );


            columns =
                this.sortColumnsByFieldState(
                    columns,
                    columnFieldIds
                );


            const rowHeaders =
                rowIndexes.map(
                    index =>
                        metadataColumns[
                            index
                        ].displayName
                );


            const columnHeaders =
                columnIndexes.map(
                    index =>
                        metadataColumns[
                            index
                        ].displayName
                );


            const categoryHeader =
                metadataColumns[
                    statusIndex
                ].displayName;


            this.renderMatrix(
                rows,
                columns,
                cells,
                rowHeaders,
                rowFieldIds,
                columnHeaders,
                columnFieldIds,
                categoryHeader
            );


            this.events.renderingFinished(
                options
            );

        }
        catch (error) {

            console.error(
                "Multi Mark Matrix error:",
                error
            );


            this.events.renderingFailed(
                options,
                String(error)
            );
        }
    }


    public getFormattingModel():
        powerbi.visuals.FormattingModel {

        return this
            .formattingSettingsService
            .buildFormattingModel(
                this.formattingSettings
            );
    }


    private renderMatrix(
        rows: MatrixAxisItem[],
        columns: MatrixAxisItem[],
        cells: Map<string, MarkPoint[]>,
        rowHeaders: string[],
        rowFieldIds: string[],
        columnHeaders: string[],
        columnFieldIds: string[],
        categoryHeader: string
    ): void {


        const appearance =
            this.formattingSettings
                .matrixAppearance;


        const valueFontFamily =
            appearance
                .valueFontFamily
                .value;


        const valueFontSize =
            appearance
                .valueFontSize
                .value;


        const headerFontFamily =
            appearance
                .headerFontFamily
                .value;


        const headerFontSize =
            appearance
                .headerFontSize
                .value;


        const valueHeaderFontFamily =
            appearance
                .valueHeaderFontFamily
                .value;


        const valueHeaderFontSize =
            appearance
                .valueHeaderFontSize
                .value;


        const rowHeight =
            appearance
                .rowHeight
                .value;


        const markGap =
            appearance
                .markGap
                .value;


        const showGrid =
            this.isHighContrast
                ? true
                : appearance
                    .showGrid
                    .value;


        const gridColor =
            this.isHighContrast
                ? this.highContrastForeground
                : appearance
                    .gridColor
                    .value.value;


        const headerBackground =
            this.isHighContrast
                ? this.highContrastBackground
                : appearance
                    .headerBackground
                    .value.value;


        const headerTextColor =
            this.isHighContrast
                ? this.highContrastForeground
                : appearance
                    .headerTextColor
                    .value.value;


        const rowHeaderBold =
            appearance
                .rowHeaderBold
                .value;


        const autoWidth =
            appearance
                .autoWidth
                .value;


        const wordWrap =
            appearance
                .wordWrap
                .value;


        const showColumnSortButtons =
            this.formattingSettings
                .sortingControls
                .showColumnSortButtons
                .value;


        const showRowSortButtons =
            this.formattingSettings
                .sortingControls
                .showRowSortButtons
                .value;


        let draggedRowGroup:
            {
                level:
                    number;
                fieldId:
                    string;
                parentKey:
                    string;
                valueKey:
                    string;
            } | null =
            null;


        const defaultRowWidth =
            160;


        const defaultDataWidth =
            100;


        const minimumColumnWidth =
            55;


        const maximumColumnWidth =
            600;


        /*
         * Auto-width caps keep extremely long values from making
         * the matrix unusably wide. Manual dragging can still make
         * a column wider, up to maximumColumnWidth.
         */
        const maximumAutoRowWidth =
            280;


        const maximumAutoDataWidth =
            240;


        const markSize =
            appearance
                .markSize
                .value;


        const headerRowHeight =
            Math.max(
                30,
                Math.max(
                    headerFontSize,
                    valueHeaderFontSize
                ) + 18
            );


        /*
         * Measure text with the same font used by the visual.
         */
        const measureCanvas =
            document.createElement(
                "canvas"
            );


        const measureContext =
            measureCanvas.getContext(
                "2d"
            );


        const measureTextWidth =
            (
                value: string,
                fontSize: number,
                fontWeight:
                    number =
                    400,
                fontFamily:
                    string =
                    valueFontFamily
            ): number => {


                const textValue =
                    value ?? "";


                if (
                    !measureContext
                ) {

                    return Math.ceil(
                        textValue.length *
                        fontSize *
                        0.6
                    );
                }


                measureContext.font =
                    `${fontWeight} ${fontSize}px ${fontFamily}`;


                return Math.ceil(
                    measureContext
                        .measureText(
                            textValue
                        )
                        .width
                );
            };


        const clampWidth =
            (
                width: number,
                maxWidth: number
            ): number =>
                Math.min(
                    maxWidth,
                    Math.max(
                        minimumColumnWidth,
                        Math.ceil(
                            width
                        )
                    )
                );


        /*
         * Auto width for the frozen Row hierarchy columns.
         * Includes both field names and actual row values.
         */
        const autoRowWidths =
            rowHeaders.map(
                (
                    header,
                    level
                ) => {


                    let measuredWidth =
                        measureTextWidth(
                            header,
                            headerFontSize,
                            600,
                            headerFontFamily
                        ) +
                        28;


                    rows.forEach(
                        row => {


                            const value =
                                row.values[
                                    level
                                ] ?? "";


                            const weight =
                                rowHeaderBold &&
                                level <
                                rowHeaders.length - 1
                                    ? 600
                                    : 400;


                            measuredWidth =
                                Math.max(
                                    measuredWidth,
                                    measureTextWidth(
                                        value,
                                        valueFontSize,
                                        weight,
                                        valueFontFamily
                                    ) +
                                    24
                                );
                        }
                    );


                    return clampWidth(
                        measuredWidth,
                        maximumAutoRowWidth
                    );
                }
            );


        /*
         * Auto width for the data/status columns.
         * Considers hierarchy header text and how many marks
         * need to fit inside a cell.
         */
        const autoDataWidths =
            columns.map(
                column => {


                    let measuredWidth =
                        minimumColumnWidth;


                    column.values.forEach(
                        value => {


                            measuredWidth =
                                Math.max(
                                    measuredWidth,
                                    measureTextWidth(
                                        value ?? "",
                                        valueHeaderFontSize,
                                        600,
                                        valueHeaderFontFamily
                                    ) +
                                    24
                                );
                        }
                    );


                    let maximumMarks =
                        0;


                    rows.forEach(
                        row => {


                            const cellKey =
                                row.key +
                                "___CELL___" +
                                column.key;


                            const marks =
                                cells.get(
                                    cellKey
                                ) ?? [];


                            maximumMarks =
                                Math.max(
                                    maximumMarks,
                                    marks.length
                                );
                        }
                    );


                    if (
                        maximumMarks > 0
                    ) {


                        const marksWidth =
                            (
                                maximumMarks *
                                markSize
                            ) +
                            (
                                Math.max(
                                    0,
                                    maximumMarks - 1
                                ) *
                                markGap
                            ) +
                            16;


                        measuredWidth =
                            Math.max(
                                measuredWidth,
                                marksWidth
                            );
                    }


                    return clampWidth(
                        measuredWidth,
                        maximumAutoDataWidth
                    );
                }
            );


        /*
         * AUTO columns follow the content when data changes.
         * Manually resized columns keep the user's width.
         */
        rowHeaders.forEach(
            (
                _header,
                level
            ) => {


                if (
                    !this
                        .manuallyResizedRowColumns
                        .has(
                            level
                        ) ||
                    this.rowColumnWidths[
                        level
                    ] === undefined
                ) {

                    this.rowColumnWidths[
                        level
                    ] =
                        autoWidth
                            ? (
                                autoRowWidths[
                                    level
                                ] ??
                                defaultRowWidth
                            )
                            : defaultRowWidth;
                }
            }
        );


        columns.forEach(
            (
                column,
                columnIndex
            ) => {


                if (
                    !this
                        .manuallyResizedDataColumns
                        .has(
                            column.key
                        ) ||
                    !this.dataColumnWidths.has(
                        column.key
                    )
                ) {

                    this.dataColumnWidths.set(
                        column.key,
                        autoWidth
                            ? (
                                autoDataWidths[
                                    columnIndex
                                ] ??
                                defaultDataWidth
                            )
                            : defaultDataWidth
                    );
                }
            }
        );


        const scroll =
            document.createElement(
                "div"
            );


        scroll.style.width =
            "100%";


        scroll.style.height =
            "100%";


        scroll.style.overflow =
            "auto";


        scroll.style.position =
            "relative";


        const table =
            document.createElement(
                "table"
            );


        table.style.borderCollapse =
            "separate";


        table.style.borderSpacing =
            "0";


        table.style.whiteSpace =
            wordWrap
                ? "normal"
                : "nowrap";


        table.style.fontFamily =
            valueFontFamily;


        table.style.fontSize =
            `${valueFontSize}px`;


        table.style.color =
            this.isHighContrast
                ? this.highContrastForeground
                : "";


        table.style.background =
            this.isHighContrast
                ? this.highContrastBackground
                : "";


        table.style.tableLayout =
            "fixed";


        const colGroup =
            document.createElement(
                "colgroup"
            );


        rowHeaders.forEach(
            (
                _header,
                level
            ) => {


                const col =
                    document.createElement(
                        "col"
                    );


                col.dataset.rowColumnLevel =
                    String(level);


                colGroup.appendChild(
                    col
                );
            }
        );


        columns.forEach(
            (
                column,
                columnIndex
            ) => {


                const col =
                    document.createElement(
                        "col"
                    );


                col.dataset.dataColumnIndex =
                    String(columnIndex);


                col.dataset.columnKey =
                    column.key;


                colGroup.appendChild(
                    col
                );
            }
        );


        table.appendChild(
            colGroup
        );


        const applyColumnWidths =
            (): void => {


                const rowWidths =
                    rowHeaders.map(
                        (
                            _header,
                            level
                        ) => {


                            const width =
                                this.rowColumnWidths[
                                    level
                                ] ??
                                defaultRowWidth;


                            return Math.min(
                                maximumColumnWidth,
                                Math.max(
                                    minimumColumnWidth,
                                    width
                                )
                            );
                        }
                    );


                const dataWidths =
                    columns.map(
                        column => {


                            const width =
                                this.dataColumnWidths.get(
                                    column.key
                                ) ??
                                defaultDataWidth;


                            return Math.min(
                                maximumColumnWidth,
                                Math.max(
                                    minimumColumnWidth,
                                    width
                                )
                            );
                        }
                    );


                const rowOffsets:
                    number[] =
                    [];


                let runningLeft =
                    0;


                rowWidths.forEach(
                    width => {

                        rowOffsets.push(
                            runningLeft
                        );


                        runningLeft +=
                            width;
                    }
                );


                const totalRowWidth =
                    rowWidths.reduce(
                        (
                            sum,
                            width
                        ) =>
                            sum +
                            width,
                        0
                    );


                const totalDataWidth =
                    dataWidths.reduce(
                        (
                            sum,
                            width
                        ) =>
                            sum +
                            width,
                        0
                    );


                table.style.width =
                    `${
                        totalRowWidth +
                        totalDataWidth
                    }px`;


                table.style.minWidth =
                    table.style.width;


                table
                    .querySelectorAll<HTMLTableColElement>(
                        "col[data-row-column-level]"
                    )
                    .forEach(
                        col => {


                            const level =
                                Number(
                                    col.dataset
                                        .rowColumnLevel
                                );


                            const width =
                                rowWidths[
                                    level
                                ];


                            col.style.width =
                                `${width}px`;
                        }
                    );


                table
                    .querySelectorAll<HTMLTableColElement>(
                        "col[data-data-column-index]"
                    )
                    .forEach(
                        col => {


                            const index =
                                Number(
                                    col.dataset
                                        .dataColumnIndex
                                );


                            const width =
                                dataWidths[
                                    index
                                ];


                            col.style.width =
                                `${width}px`;
                        }
                    );


                table
                    .querySelectorAll<HTMLElement>(
                        "[data-row-level]"
                    )
                    .forEach(
                        element => {


                            const level =
                                Number(
                                    element.dataset
                                        .rowLevel
                                );


                            const width =
                                rowWidths[
                                    level
                                ];


                            const left =
                                rowOffsets[
                                    level
                                ];


                            element.style.width =
                                `${width}px`;


                            element.style.minWidth =
                                `${width}px`;


                            element.style.maxWidth =
                                `${width}px`;


                            element.style.left =
                                `${left}px`;
                        }
                    );


                table
                    .querySelectorAll<HTMLElement>(
                        "[data-hierarchy-blank]"
                    )
                    .forEach(
                        element => {


                            const blankWidth =
                                rowWidths
                                    .slice(
                                        0,
                                        Math.max(
                                            0,
                                            rowWidths.length -
                                            1
                                        )
                                    )
                                    .reduce(
                                        (
                                            sum,
                                            width
                                        ) =>
                                            sum +
                                            width,
                                        0
                                    );


                            element.style.width =
                                `${blankWidth}px`;


                            element.style.minWidth =
                                `${blankWidth}px`;


                            element.style.maxWidth =
                                `${blankWidth}px`;


                            element.style.left =
                                "0px";
                        }
                    );


                table
                    .querySelectorAll<HTMLElement>(
                        "[data-hierarchy-name]"
                    )
                    .forEach(
                        element => {


                            const lastLevel =
                                Math.max(
                                    0,
                                    rowWidths.length -
                                    1
                                );


                            const width =
                                rowWidths[
                                    lastLevel
                                ] ??
                                defaultRowWidth;


                            const left =
                                rowOffsets[
                                    lastLevel
                                ] ??
                                0;


                            element.style.width =
                                `${width}px`;


                            element.style.minWidth =
                                `${width}px`;


                            element.style.maxWidth =
                                `${width}px`;


                            element.style.left =
                                `${left}px`;
                        }
                    );


                table
                    .querySelectorAll<HTMLElement>(
                        "[data-data-cell-index]"
                    )
                    .forEach(
                        element => {


                            const index =
                                Number(
                                    element.dataset
                                        .dataCellIndex
                                );


                            const width =
                                dataWidths[
                                    index
                                ];


                            element.style.width =
                                `${width}px`;


                            element.style.minWidth =
                                `${width}px`;


                            element.style.maxWidth =
                                `${width}px`;
                        }
                    );


                /*
                 * Header rows can become taller when text wraps.
                 * Recalculate sticky top offsets from the REAL
                 * rendered row heights, so sticky headers never overlap.
                 */
                let stickyTopOffset =
                    0;


                thead
                    .querySelectorAll<HTMLTableRowElement>(
                        "tr"
                    )
                    .forEach(
                        headerRow => {


                            headerRow
                                .querySelectorAll<HTMLElement>(
                                    "th"
                                )
                                .forEach(
                                    headerCell => {

                                        headerCell.style.top =
                                            `${stickyTopOffset}px`;
                                    }
                                );


                            stickyTopOffset +=
                                headerRow
                                    .getBoundingClientRect()
                                    .height;
                        }
                    );
            };


        const thead =
            document.createElement(
                "thead"
            );


        columnHeaders.forEach(
            (
                columnHeader,
                level
            ) => {


                const tr =
                    document.createElement(
                        "tr"
                    );


                tr.style.height =
                    `${headerRowHeight}px`;


                const stickyTop =
                    level *
                    headerRowHeight;


                if (
                    rowHeaders.length > 1
                ) {


                    const blank =
                        document.createElement(
                            "th"
                        );


                    blank.colSpan =
                        rowHeaders.length - 1;


                    blank.textContent =
                        "";


                    blank.dataset.hierarchyBlank =
                        "true";


                    this.styleHeader(
                        blank,
                        headerBackground,
                        headerTextColor,
                        showGrid,
                        gridColor,
                        headerFontSize
                    );


                    blank.style.position =
                        "sticky";


                    blank.style.top =
                        `${stickyTop}px`;


                    blank.style.zIndex =
                        "8";


                    blank.style.boxSizing =
                        "border-box";


                    tr.appendChild(
                        blank
                    );
                }


                const hierarchyName =
                    document.createElement(
                        "th"
                    );


                const hierarchyHeaderContent =
                    document.createElement(
                        "div"
                    );


                hierarchyHeaderContent.style.display =
                    "flex";


                hierarchyHeaderContent.style.alignItems =
                    "center";


                hierarchyHeaderContent.style.gap =
                    "6px";


                hierarchyHeaderContent.style.minWidth =
                    "0";


                const hierarchyHeaderLabel =
                    document.createElement(
                        "span"
                    );


                hierarchyHeaderLabel.textContent =
                    columnHeader;


                hierarchyHeaderLabel.style.flex =
                    "1 1 auto";


                hierarchyHeaderLabel.style.minWidth =
                    "0";


                hierarchyHeaderLabel.style.overflow =
                    "hidden";


                hierarchyHeaderLabel.style.textOverflow =
                    "ellipsis";


                const columnFieldId =
                    columnFieldIds[
                        level
                    ] ??
                    `columns:${level}:${columnHeader}`;


                const sortMode =
                    this.columnFieldSortState.get(
                        columnFieldId
                    ) ??
                    "default";


                const sortButton =
                    document.createElement(
                        "button"
                    );


                sortButton.type =
                    "button";


                sortButton.textContent =
                    sortMode ===
                        "ascending"
                        ? "↑"
                        : sortMode ===
                            "descending"
                            ? "↓"
                            : "↕";


                sortButton.title =
                    sortMode ===
                        "ascending"
                        ? `Sort ${columnHeader}: Ascending. Click for Descending.`
                        : sortMode ===
                            "descending"
                            ? `Sort ${columnHeader}: Descending. Click for Data order.`
                            : `Sort ${columnHeader}: Data order. Click for Ascending.`;


                sortButton.setAttribute(
                    "aria-label",
                    sortButton.title
                );


                sortButton.style.flex =
                    "0 0 auto";


                sortButton.style.width =
                    "22px";


                sortButton.style.height =
                    "22px";


                sortButton.style.padding =
                    "0";


                sortButton.style.borderRadius =
                    "3px";


                sortButton.style.border =
                    `1px solid ${
                        this.isHighContrast
                            ? this.highContrastForeground
                            : "#B3B3B3"
                    }`;


                sortButton.style.background =
                    this.isHighContrast
                        ? this.highContrastBackground
                        : "#FFFFFF";


                sortButton.style.color =
                    this.isHighContrast
                        ? this.highContrastForeground
                        : headerTextColor;


                sortButton.style.cursor =
                    this.allowInteractions
                        ? "pointer"
                        : "default";


                sortButton.disabled =
                    !this.allowInteractions;


                sortButton.addEventListener(
                    "click",
                    (
                        event:
                            MouseEvent
                    ) => {


                        event.preventDefault();
                        event.stopPropagation();


                        if (
                            !this.allowInteractions
                        ) {

                            return;
                        }


                        const currentMode =
                            this.columnFieldSortState.get(
                                columnFieldId
                            ) ??
                            "default";


                        if (
                            currentMode ===
                                "default"
                        ) {

                            this.columnFieldSortState.set(
                                columnFieldId,
                                "ascending"
                            );
                        }
                        else if (
                            currentMode ===
                                "ascending"
                        ) {

                            this.columnFieldSortState.set(
                                columnFieldId,
                                "descending"
                            );
                        }
                        else {

                            this.columnFieldSortState.delete(
                                columnFieldId
                            );
                        }


                        this.persistColumnFieldSortState();
                    }
                );


                hierarchyHeaderContent.appendChild(
                    hierarchyHeaderLabel
                );


                if (
                    showColumnSortButtons
                ) {

                    hierarchyHeaderContent.appendChild(
                        sortButton
                    );
                }


                hierarchyName.appendChild(
                    hierarchyHeaderContent
                );


                hierarchyName.dataset.hierarchyName =
                    "true";


                hierarchyName.style.textAlign =
                    "left";


                hierarchyName.style.whiteSpace =
                    wordWrap
                        ? "normal"
                        : "nowrap";


                hierarchyName.style.overflowWrap =
                    wordWrap
                        ? "anywhere"
                        : "normal";


                hierarchyName.style.wordBreak =
                    wordWrap
                        ? "break-word"
                        : "normal";


                hierarchyName.style.overflow =
                    "hidden";


                hierarchyName.style.textOverflow =
                    wordWrap
                        ? ""
                        : "ellipsis";


                hierarchyName.style.lineHeight =
                    "1.25";


                hierarchyName.style.boxSizing =
                    "border-box";


                this.styleHeader(
                    hierarchyName,
                    headerBackground,
                    headerTextColor,
                    showGrid,
                    gridColor,
                    headerFontSize
                );


                hierarchyName.style.fontFamily =
                    headerFontFamily;


                hierarchyName.style.position =
                    "sticky";


                hierarchyName.style.top =
                    `${stickyTop}px`;


                hierarchyName.style.zIndex =
                    "9";


                tr.appendChild(
                    hierarchyName
                );


                let columnIndex =
                    0;


                while (
                    columnIndex <
                    columns.length
                ) {


                    const column =
                        columns[
                            columnIndex
                        ];


                    const span =
                        this.getColumnSpan(
                            columns,
                            columnIndex,
                            level
                        );


                    const th =
                        document.createElement(
                            "th"
                        );


                    th.textContent =
                        column.values[
                            level
                        ] ?? "";


                    th.colSpan =
                        span;


                    th.style.textAlign =
                        "center";


                    th.style.whiteSpace =
                        wordWrap
                            ? "normal"
                            : "nowrap";


                    th.style.overflowWrap =
                        wordWrap
                            ? "anywhere"
                            : "normal";


                    th.style.wordBreak =
                        wordWrap
                            ? "break-word"
                            : "normal";


                    th.style.overflow =
                        "hidden";


                    th.style.textOverflow =
                        wordWrap
                            ? ""
                            : "ellipsis";


                    th.style.lineHeight =
                        "1.25";


                    this.styleHeader(
                        th,
                        headerBackground,
                        headerTextColor,
                        showGrid,
                        gridColor,
                        valueHeaderFontSize
                    );


                    th.style.fontFamily =
                        valueHeaderFontFamily;


                    th.style.position =
                        "sticky";


                    th.style.top =
                        `${stickyTop}px`;


                    th.style.zIndex =
                        "5";


                    if (
                        level ===
                        columnHeaders.length -
                        1 &&
                        span === 1
                    ) {


                        th.dataset.dataCellIndex =
                            String(
                                columnIndex
                            );


                        const currentColumnKey =
                            column.key;


                        this.addColumnResizeHandle(
                            th,

                            () =>
                                this
                                    .dataColumnWidths
                                    .get(
                                        currentColumnKey
                                    ) ??
                                defaultDataWidth,

                            (
                                newWidth,
                                isManual
                            ) => {


                                this
                                    .dataColumnWidths
                                    .set(
                                        currentColumnKey,
                                        newWidth
                                    );


                                if (
                                    isManual
                                ) {

                                    this
                                        .manuallyResizedDataColumns
                                        .add(
                                            currentColumnKey
                                        );
                                }
                                else {

                                    this
                                        .manuallyResizedDataColumns
                                        .delete(
                                            currentColumnKey
                                        );
                                }


                                applyColumnWidths();
                            },

                            autoWidth
                                ? (
                                    autoDataWidths[
                                        columnIndex
                                    ] ??
                                    defaultDataWidth
                                )
                                : defaultDataWidth,

                            minimumColumnWidth,

                            maximumColumnWidth
                        );
                    }


                    tr.appendChild(
                        th
                    );


                    columnIndex +=
                        span;
                }


                thead.appendChild(
                    tr
                );
            }
        );


        const rowHeaderRow =
            document.createElement(
                "tr"
            );


        rowHeaderRow.style.height =
            `${headerRowHeight}px`;


        const rowHeaderTop =
            columnHeaders.length *
            headerRowHeight;


        rowHeaders.forEach(
            (
                header,
                level
            ) => {


                const th =
                    document.createElement(
                        "th"
                    );


                const rowHeaderContent =
                    document.createElement(
                        "div"
                    );


                rowHeaderContent.style.display =
                    "flex";


                rowHeaderContent.style.alignItems =
                    "center";


                rowHeaderContent.style.gap =
                    "6px";


                rowHeaderContent.style.minWidth =
                    "0";


                const rowHeaderLabel =
                    document.createElement(
                        "span"
                    );


                rowHeaderLabel.textContent =
                    header;


                rowHeaderLabel.style.flex =
                    "1 1 auto";


                rowHeaderLabel.style.minWidth =
                    "0";


                rowHeaderLabel.style.overflow =
                    "hidden";


                rowHeaderLabel.style.textOverflow =
                    "ellipsis";


                const rowFieldId =
                    rowFieldIds[
                        level
                    ] ??
                    `rows:${level}:${header}`;


                const rowSortMode =
                    this.rowFieldSortState.get(
                        rowFieldId
                    ) ??
                    "default";


                const rowSortButton =
                    document.createElement(
                        "button"
                    );


                rowSortButton.type =
                    "button";


                rowSortButton.textContent =
                    rowSortMode ===
                        "ascending"
                        ? "↑"
                        : rowSortMode ===
                            "descending"
                            ? "↓"
                            : rowSortMode ===
                                "manual"
                                ? "≡"
                                : "↕";


                rowSortButton.title =
                    rowSortMode ===
                        "ascending"
                        ? `Sort ${header}: Ascending. Click for Descending.`
                        : rowSortMode ===
                            "descending"
                            ? `Sort ${header}: Descending. Click for Manual.`
                            : rowSortMode ===
                                "manual"
                                ? `Sort ${header}: Manual. Drag row values to reorder. Click for Data order.`
                                : `Sort ${header}: Data order. Click for Ascending.`;


                rowSortButton.setAttribute(
                    "aria-label",
                    rowSortButton.title
                );


                rowSortButton.style.flex =
                    "0 0 auto";


                rowSortButton.style.width =
                    "22px";


                rowSortButton.style.height =
                    "22px";


                rowSortButton.style.padding =
                    "0";


                rowSortButton.style.borderRadius =
                    "3px";


                rowSortButton.style.border =
                    `1px solid ${
                        this.isHighContrast
                            ? this.highContrastForeground
                            : "#B3B3B3"
                    }`;


                rowSortButton.style.background =
                    this.isHighContrast
                        ? this.highContrastBackground
                        : "#FFFFFF";


                rowSortButton.style.color =
                    this.isHighContrast
                        ? this.highContrastForeground
                        : headerTextColor;


                rowSortButton.style.cursor =
                    this.allowInteractions
                        ? "pointer"
                        : "default";


                rowSortButton.disabled =
                    !this.allowInteractions;


                rowSortButton.addEventListener(
                    "click",
                    (
                        event:
                            MouseEvent
                    ) => {


                        event.preventDefault();
                        event.stopPropagation();


                        if (
                            !this.allowInteractions
                        ) {

                            return;
                        }


                        const currentMode =
                            this.rowFieldSortState.get(
                                rowFieldId
                            ) ??
                            "default";


                        if (
                            currentMode ===
                                "default"
                        ) {

                            this.rowFieldSortState.set(
                                rowFieldId,
                                "ascending"
                            );
                        }
                        else if (
                            currentMode ===
                                "ascending"
                        ) {

                            this.rowFieldSortState.set(
                                rowFieldId,
                                "descending"
                            );
                        }
                        else if (
                            currentMode ===
                                "descending"
                        ) {

                            this.rowFieldSortState.set(
                                rowFieldId,
                                "manual"
                            );
                        }
                        else {

                            this.rowFieldSortState.delete(
                                rowFieldId
                            );
                        }


                        this.persistRowSortState();
                    }
                );


                rowHeaderContent.appendChild(
                    rowHeaderLabel
                );


                if (
                    showRowSortButtons
                ) {

                    rowHeaderContent.appendChild(
                        rowSortButton
                    );
                }


                th.appendChild(
                    rowHeaderContent
                );


                th.dataset.rowLevel =
                    String(level);


                th.style.textAlign =
                    "left";


                th.style.whiteSpace =
                    wordWrap
                        ? "normal"
                        : "nowrap";


                th.style.overflowWrap =
                    wordWrap
                        ? "anywhere"
                        : "normal";


                th.style.wordBreak =
                    wordWrap
                        ? "break-word"
                        : "normal";


                th.style.overflow =
                    "hidden";


                th.style.textOverflow =
                    wordWrap
                        ? ""
                        : "ellipsis";


                th.style.lineHeight =
                    "1.25";


                th.style.boxSizing =
                    "border-box";


                this.styleHeader(
                    th,
                    headerBackground,
                    headerTextColor,
                    showGrid,
                    gridColor,
                    headerFontSize
                );


                th.style.fontFamily =
                    headerFontFamily;


                th.style.position =
                    "sticky";


                th.style.top =
                    `${rowHeaderTop}px`;


                th.style.zIndex =
                    "10";


                const currentLevel =
                    level;


                this.addColumnResizeHandle(
                    th,

                    () =>
                        this.rowColumnWidths[
                            currentLevel
                        ] ??
                        defaultRowWidth,

                    (
                        newWidth,
                        isManual
                    ) => {


                        this.rowColumnWidths[
                            currentLevel
                        ] =
                            newWidth;


                        if (
                            isManual
                        ) {

                            this
                                .manuallyResizedRowColumns
                                .add(
                                    currentLevel
                                );
                        }
                        else {

                            this
                                .manuallyResizedRowColumns
                                .delete(
                                    currentLevel
                                );
                        }


                        applyColumnWidths();
                    },

                    autoWidth
                        ? (
                            autoRowWidths[
                                currentLevel
                            ] ??
                            defaultRowWidth
                        )
                        : defaultRowWidth,

                    minimumColumnWidth,

                    maximumColumnWidth
                );


                rowHeaderRow.appendChild(
                    th
                );
            }
        );


        columns.forEach(
            (
                _column,
                columnIndex
            ) => {


                const th =
                    document.createElement(
                        "th"
                    );


                th.textContent =
                    "";


                th.dataset.dataCellIndex =
                    String(
                        columnIndex
                    );


                this.styleHeader(
                    th,
                    headerBackground,
                    headerTextColor,
                    showGrid,
                    gridColor,
                    headerFontSize
                );


                th.style.position =
                    "sticky";


                th.style.top =
                    `${rowHeaderTop}px`;


                th.style.zIndex =
                    "5";


                rowHeaderRow.appendChild(
                    th
                );
            }
        );


        thead.appendChild(
            rowHeaderRow
        );


        table.appendChild(
            thead
        );


        const tbody =
            document.createElement(
                "tbody"
            );


        if (
            rows.length ===
                0
        ) {


            const emptyRow =
                document.createElement(
                    "tr"
                );


            const emptyCell =
                document.createElement(
                    "td"
                );


            emptyCell.colSpan =
                Math.max(
                    1,
                    rowHeaders.length +
                    columns.length
                );


            emptyCell.textContent =
                "No data for the current filters.";


            emptyCell.setAttribute(
                "role",
                "status"
            );


            emptyCell.setAttribute(
                "aria-live",
                "polite"
            );


            emptyCell.style.height =
                `${Math.max(
                    rowHeight * 2,
                    52
                )}px`;


            emptyCell.style.padding =
                "14px 16px";


            emptyCell.style.textAlign =
                "center";


            emptyCell.style.verticalAlign =
                "middle";


            emptyCell.style.fontSize =
                `${valueFontSize}px`;


            emptyCell.style.fontStyle =
                "italic";


            emptyCell.style.boxSizing =
                "border-box";


            emptyCell.style.color =
                this.isHighContrast
                    ? this.highContrastForeground
                    : "#666666";


            emptyCell.style.background =
                this.isHighContrast
                    ? this.highContrastBackground
                    : "#FFFFFF";


            if (
                showGrid
            ) {

                emptyCell.style.borderBottom =
                    `1px solid ${gridColor}`;
            }


            emptyRow.appendChild(
                emptyCell
            );


            tbody.appendChild(
                emptyRow
            );
        }


        rows.forEach(
            (
                row,
                rowIndex
            ) => {


                const tr =
                    document.createElement(
                        "tr"
                    );


                row.values.forEach(
                    (
                        value,
                        level
                    ) => {


                        if (
                            !this.shouldRenderRowCell(
                                rows,
                                rowIndex,
                                level
                            )
                        ) {

                            return;
                        }


                        const span =
                            this.getRowSpan(
                                rows,
                                rowIndex,
                                level
                            );


                        const td =
                            document.createElement(
                                "td"
                            );


                        td.textContent =
                            value;


                        td.rowSpan =
                            span;


                        td.dataset.rowLevel =
                            String(level);


                        td.style.padding =
                            "4px 8px";


                        td.style.height =
                            `${rowHeight}px`;


                        td.style.verticalAlign =
                            "top";


                        td.style.boxSizing =
                            "border-box";


                        td.style.overflow =
                            "hidden";


                        td.style.whiteSpace =
                            wordWrap
                                ? "normal"
                                : "nowrap";


                        td.style.overflowWrap =
                            wordWrap
                                ? "anywhere"
                                : "normal";


                        td.style.wordBreak =
                            wordWrap
                                ? "break-word"
                                : "normal";


                        td.style.textOverflow =
                            wordWrap
                                ? ""
                                : "ellipsis";


                        td.style.lineHeight =
                            "1.25";


                        td.style.position =
                            "sticky";


                        td.style.zIndex =
                            "3";


                        td.style.background =
                            this.isHighContrast
                                ? this.highContrastBackground
                                : "#FFFFFF";


                        td.style.color =
                            this.isHighContrast
                                ? this.highContrastForeground
                                : "";


                        if (
                            showGrid
                        ) {

                            td.style.borderBottom =
                                `1px solid ${gridColor}`;


                            td.style.borderRight =
                                `1px solid ${gridColor}`;
                        }


                        if (
                            rowHeaderBold &&
                            level <
                            row.values.length - 1
                        ) {

                            td.style.fontWeight =
                                "600";
                        }


                        const rowFieldId =
                            rowFieldIds[
                                level
                            ] ??
                            `rows:${level}:${
                                rowHeaders[
                                    level
                                ] ?? ""
                            }`;


                        const rowSortMode =
                            this.rowFieldSortState.get(
                                rowFieldId
                            ) ??
                            "default";


                        if (
                            rowSortMode ===
                                "manual" &&
                            showRowSortButtons
                        ) {


                            const parentKey =
                                this.getAxisParentKey(
                                    row,
                                    level
                                );


                            const valueKey =
                                this.getAxisValueKey(
                                    row,
                                    level
                                );


                            td.draggable =
                                true;


                            td.style.cursor =
                                "grab";


                            td.title =
                                "Drag to change row order within this group.";


                            td.addEventListener(
                                "dragstart",
                                (
                                    event:
                                        DragEvent
                                ) => {


                                    draggedRowGroup = {
                                        level,
                                        fieldId:
                                            rowFieldId,
                                        parentKey,
                                        valueKey
                                    };


                                    td.style.opacity =
                                        "0.55";


                                    if (
                                        event.dataTransfer
                                    ) {

                                        event
                                            .dataTransfer
                                            .effectAllowed =
                                            "move";
                                    }
                                }
                            );


                            td.addEventListener(
                                "dragend",
                                () => {


                                    draggedRowGroup =
                                        null;


                                    td.style.opacity =
                                        "1";
                                }
                            );


                            td.addEventListener(
                                "dragover",
                                (
                                    event:
                                        DragEvent
                                ) => {


                                    if (
                                        !draggedRowGroup ||
                                        draggedRowGroup.level !==
                                            level ||
                                        draggedRowGroup.fieldId !==
                                            rowFieldId ||
                                        draggedRowGroup.parentKey !==
                                            parentKey ||
                                        draggedRowGroup.valueKey ===
                                            valueKey
                                    ) {

                                        return;
                                    }


                                    event.preventDefault();


                                    if (
                                        event.dataTransfer
                                    ) {

                                        event
                                            .dataTransfer
                                            .dropEffect =
                                            "move";
                                    }


                                    td.style.outline =
                                        `2px dashed ${
                                            this.isHighContrast
                                                ? this.highContrastForeground
                                                : "#666666"
                                        }`;


                                    td.style.outlineOffset =
                                        "-2px";
                                }
                            );


                            td.addEventListener(
                                "dragleave",
                                () => {


                                    td.style.outline =
                                        "";


                                    td.style.outlineOffset =
                                        "";
                                }
                            );


                            td.addEventListener(
                                "drop",
                                (
                                    event:
                                        DragEvent
                                ) => {


                                    event.preventDefault();


                                    td.style.outline =
                                        "";


                                    td.style.outlineOffset =
                                        "";


                                    if (
                                        !draggedRowGroup ||
                                        draggedRowGroup.level !==
                                            level ||
                                        draggedRowGroup.fieldId !==
                                            rowFieldId ||
                                        draggedRowGroup.parentKey !==
                                            parentKey ||
                                        draggedRowGroup.valueKey ===
                                            valueKey
                                    ) {

                                        return;
                                    }


                                    const siblingKeys:
                                        string[] =
                                        [];


                                    const seen =
                                        new Set<string>();


                                    rows.forEach(
                                        rowItem => {


                                            if (
                                                this.getAxisParentKey(
                                                    rowItem,
                                                    level
                                                ) !==
                                                    parentKey
                                            ) {

                                                return;
                                            }


                                            const siblingKey =
                                                this.getAxisValueKey(
                                                    rowItem,
                                                    level
                                                );


                                            if (
                                                seen.has(
                                                    siblingKey
                                                )
                                            ) {

                                                return;
                                            }


                                            seen.add(
                                                siblingKey
                                            );


                                            siblingKeys.push(
                                                siblingKey
                                            );
                                        }
                                    );


                                    const sourceIndex =
                                        siblingKeys.indexOf(
                                            draggedRowGroup
                                                .valueKey
                                        );


                                    const targetIndex =
                                        siblingKeys.indexOf(
                                            valueKey
                                        );


                                    if (
                                        sourceIndex <
                                            0 ||
                                        targetIndex <
                                            0
                                    ) {

                                        return;
                                    }


                                    const reordered =
                                        [
                                            ...siblingKeys
                                        ];


                                    reordered.splice(
                                        sourceIndex,
                                        1
                                    );


                                    let insertionIndex =
                                        reordered.indexOf(
                                            valueKey
                                        );


                                    if (
                                        insertionIndex <
                                            0
                                    ) {

                                        return;
                                    }


                                    const rect =
                                        td
                                            .getBoundingClientRect();


                                    const insertAfter =
                                        event.clientY >
                                        (
                                            rect.top +
                                            rect.height /
                                            2
                                        );


                                    if (
                                        insertAfter
                                    ) {

                                        insertionIndex +=
                                            1;
                                    }


                                    reordered.splice(
                                        insertionIndex,
                                        0,
                                        draggedRowGroup
                                            .valueKey
                                    );


                                    draggedRowGroup =
                                        null;


                                    this.setManualRowOrder(
                                        rowFieldId,
                                        parentKey,
                                        reordered
                                    );


                                    const reorderedRows =
                                        this.sortRowsByFieldState(
                                            [
                                                ...rows
                                            ],
                                            rowFieldIds
                                        );


                                    this.target
                                        .replaceChildren();


                                    this.renderMatrix(
                                        reorderedRows,
                                        columns,
                                        cells,
                                        rowHeaders,
                                        rowFieldIds,
                                        columnHeaders,
                                        columnFieldIds,
                                        categoryHeader
                                    );
                                }
                            );
                        }


                        tr.appendChild(
                            td
                        );
                    }
                );


                columns.forEach(
                    (
                        column,
                        columnIndex
                    ) => {


                        const td =
                            document.createElement(
                                "td"
                            );


                        td.dataset.dataCellIndex =
                            String(
                                columnIndex
                            );


                        td.style.padding =
                            "4px 6px";


                        td.style.height =
                            `${rowHeight}px`;


                        td.style.boxSizing =
                            "border-box";


                        td.style.background =
                            this.isHighContrast
                                ? this.highContrastBackground
                                : "#FFFFFF";


                        td.style.color =
                            this.isHighContrast
                                ? this.highContrastForeground
                                : "";


                        if (
                            showGrid
                        ) {

                            td.style.borderLeft =
                                `1px solid ${gridColor}`;


                            td.style.borderBottom =
                                `1px solid ${gridColor}`;
                        }


                        const marksContainer =
                            document.createElement(
                                "div"
                            );


                        marksContainer.style.display =
                            "flex";


                        marksContainer.style.flexDirection =
                            "row";


                        marksContainer.style.flexWrap =
                            wordWrap
                                ? "wrap"
                                : "nowrap";


                        marksContainer.style.overflow =
                            wordWrap
                                ? "visible"
                                : "hidden";


                        marksContainer.style.alignItems =
                            "center";


                        marksContainer.style.justifyContent =
                            "flex-start";


                        marksContainer.style.gap =
                            `${markGap}px`;


                        marksContainer.style.minHeight =
                            `${rowHeight}px`;


                        const cellKey =
                            row.key +
                            "___CELL___" +
                            column.key;


                        const marks =
                            cells.get(
                                cellKey
                            ) ?? [];


                        marks.forEach(
                            mark => {


                                const markElement =
                                    this.createMark(
                                        mark
                                    );


                                marksContainer.appendChild(
                                    markElement
                                );
                            }
                        );


                        td.appendChild(
                            marksContainer
                        );


                        tr.appendChild(
                            td
                        );
                    }
                );


                tbody.appendChild(
                    tr
                );
            }
        );


        table.appendChild(
            tbody
        );


        scroll.appendChild(
            table
        );


        /*
         * Legend + matrix layout.
         * The matrix keeps its own scroll container, so sticky
         * headers and sticky Row columns continue to work.
         */
        const root =
            document.createElement(
                "div"
            );


        root.style.width =
            "100%";


        root.style.height =
            "100%";


        root.style.display =
            "flex";


        root.style.minWidth =
            "0";


        root.style.minHeight =
            "0";


        root.style.overflow =
            "hidden";


        root.style.background =
            this.isHighContrast
                ? this.highContrastBackground
                : "";


        root.style.color =
            this.isHighContrast
                ? this.highContrastForeground
                : "";


        scroll.style.flex =
            "1 1 auto";


        scroll.style.minWidth =
            "0";


        scroll.style.minHeight =
            "0";


        const legendSettings =
            this.formattingSettings
                .legend;


        const showLegend =
            legendSettings
                .showLegend
                .value;


        const legendPosition =
            String(
                legendSettings
                    .position
                    .value
                    .value
            );


        const legend =
            showLegend
                ? this.createLegend(
                    cells,
                    categoryHeader
                )
                : null;


        if (
            legendPosition ===
            "left" ||
            legendPosition ===
            "right"
        ) {

            root.style.flexDirection =
                "row";
        }
        else {

            root.style.flexDirection =
                "column";
        }


        if (
            legend &&
            (
                legendPosition ===
                "top" ||
                legendPosition ===
                "left"
            )
        ) {

            root.appendChild(
                legend
            );
        }


        root.appendChild(
            scroll
        );


        if (
            legend &&
            (
                legendPosition ===
                "bottom" ||
                legendPosition ===
                "right"
            )
        ) {

            root.appendChild(
                legend
            );
        }


        this.target.appendChild(
            root
        );


        applyColumnWidths();
    }


    /*
     * Legend uses only categories that are currently present
     * in the visual. Colors come from the same Category colors
     * settings as the marks, so legend and matrix never diverge.
     */
    private createLegend(
        cells:
            Map<string, MarkPoint[]>,
        categoryHeader:
            string
    ): HTMLElement {


        const legendSettings =
            this.formattingSettings
                .legend;


        const fontFamily =
            legendSettings
                .fontFamily
                .value;


        const fontSize =
            legendSettings
                .fontSize
                .value;


        const titleFontFamily =
            legendSettings
                .titleFontFamily
                .value;


        const titleFontSize =
            legendSettings
                .titleFontSize
                .value;


        const showTitle =
            legendSettings
                .showTitle
                .value;


        const configuredTitle =
            legendSettings
                .titleText
                .value
                .trim();


        const legendTitle =
            configuredTitle.length > 0
                ? configuredTitle
                : categoryHeader;


        const position =
            String(
                legendSettings
                    .position
                    .value
                    .value
            );


        const knownStatusOrder =
            [
                "Durchzuführende Analyse",
                "Analyse in Bearbeitung",
                "Analyse durchgeführt",
                "Analyse gereviewed",
                "Analyse freigegeben",
                "Investigation open",
                "Sample rejected/canceled",
                "Analyse rejected/canceled",
                "Analyse r/c + Batch closed",
                "Not mapped"
            ];


        const statuses =
            new Set<string>();


        cells.forEach(
            marks => {

                marks.forEach(
                    mark => {

                        if (
                            mark.status
                        ) {

                            statuses.add(
                                mark.status
                            );
                        }
                    }
                );
            }
        );


        const orderedStatuses =
            Array.from(
                statuses
            )
                .sort(
                    (
                        a,
                        b
                    ) => {


                        const indexA =
                            knownStatusOrder
                                .indexOf(
                                    a
                                );


                        const indexB =
                            knownStatusOrder
                                .indexOf(
                                    b
                                );


                        if (
                            indexA >= 0 &&
                            indexB >= 0
                        ) {

                            return (
                                indexA -
                                indexB
                            );
                        }


                        if (
                            indexA >= 0
                        ) {

                            return -1;
                        }


                        if (
                            indexB >= 0
                        ) {

                            return 1;
                        }


                        return a.localeCompare(
                            b
                        );
                    }
                );


        const legend =
            document.createElement(
                "div"
            );


        legend.setAttribute(
            "role",
            "group"
        );


        legend.setAttribute(
            "aria-label",
            `${legendTitle} legend`
        );


        legend.style.display =
            "flex";


        legend.style.boxSizing =
            "border-box";


        legend.style.fontFamily =
            fontFamily;


        legend.style.fontSize =
            `${fontSize}px`;


        legend.style.color =
            this.isHighContrast
                ? this.highContrastForeground
                : "#333333";


        legend.style.background =
            this.isHighContrast
                ? this.highContrastBackground
                : "#FFFFFF";


        legend.style.gap =
            "10px";


        legend.style.padding =
            "6px 8px";


        legend.style.flex =
            "0 0 auto";


        if (
            position ===
            "left" ||
            position ===
            "right"
        ) {

            legend.style.flexDirection =
                "column";


            legend.style.alignItems =
                "flex-start";


            legend.style.maxWidth =
                "220px";


            legend.style.overflowY =
                "auto";


            legend.style.overflowX =
                "hidden";
        }
        else {

            legend.style.flexDirection =
                "row";


            legend.style.alignItems =
                "center";


            legend.style.flexWrap =
                "wrap";


            legend.style.maxWidth =
                "100%";
        }


        if (
            showTitle
        ) {

            const title =
                document.createElement(
                    "div"
                );


            title.textContent =
                legendTitle;


            title.style.fontFamily =
                titleFontFamily;


            title.style.fontSize =
                `${titleFontSize}px`;


            title.style.fontWeight =
                "600";


            title.style.whiteSpace =
                "nowrap";


            title.style.marginRight =
                (
                    position ===
                    "top" ||
                    position ===
                    "bottom"
                )
                    ? "4px"
                    : "0";


            legend.appendChild(
                title
            );
        }


        orderedStatuses.forEach(
            status => {


                const item =
                    document.createElement(
                        "div"
                    );


                item.style.display =
                    "inline-flex";


                item.style.alignItems =
                    "center";


                item.style.gap =
                    "6px";


                item.style.minWidth =
                    "0";


                item.style.maxWidth =
                    "100%";


                const marker =
                    document.createElement(
                        "span"
                    );


                const markerSize =
                    Math.max(
                        10,
                        Math.min(
                            14,
                            fontSize + 2
                        )
                    );


                marker.style.display =
                    "inline-block";


                marker.style.width =
                    `${markerSize}px`;


                marker.style.height =
                    `${markerSize}px`;


                marker.style.minWidth =
                    `${markerSize}px`;


                marker.style.minHeight =
                    `${markerSize}px`;


                marker.style.backgroundColor =
                    this.isHighContrast
                        ? this.highContrastBackground
                        : this.getStatusColor(
                            status
                        );


                marker.style.boxSizing =
                    "border-box";


                if (
                    this.isHighContrast
                ) {

                    marker.style.border =
                        `2px solid ${this.highContrastForeground}`;
                }


                const shape =
                    this.getStatusShape(
                        status
                    );


                switch (
                    shape
                ) {


                    case "circle":

                        marker.style.borderRadius =
                            "50%";

                        break;


                    case "diamond":

                        marker.style.clipPath =
                            "polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%)";

                        break;


                    case "triangle":

                        marker.style.clipPath =
                            "polygon(50% 0%, 100% 100%, 0% 100%)";

                        break;


                    default:

                        marker.style.borderRadius =
                            "0";

                        break;
                }


                const label =
                    document.createElement(
                        "span"
                    );


                label.textContent =
                    status;


                label.style.whiteSpace =
                    (
                        position ===
                        "left" ||
                        position ===
                        "right"
                    )
                        ? "normal"
                        : "nowrap";


                label.style.overflowWrap =
                    "anywhere";


                item.appendChild(
                    marker
                );


                item.appendChild(
                    label
                );


                legend.appendChild(
                    item
                );
            }
        );


        return legend;
    }


    /*
     * Drag the right edge of a header to resize.
     * Double-click the edge to restore the default width.
     */
    /*
     * Restore Rows sort modes and manual orders from PBIX metadata.
     */
    private restoreRowSortState(
        dataView:
            powerbi.DataView
    ): void {


        const rawValue =
            dataView
                .metadata
                .objects
                ?.internalState
                ?.rowSortState;


        const serialized =
            typeof rawValue ===
                "string"
                ? rawValue
                : "";


        if (
            this.pendingRowSortState !==
                null
        ) {


            if (
                serialized ===
                this.pendingRowSortState
            ) {

                this.lastRowSortStateFromDataView =
                    serialized;


                this.pendingRowSortState =
                    null;


                return;
            }


            if (
                serialized ===
                this.lastRowSortStateFromDataView
            ) {

                return;
            }


            this.pendingRowSortState =
                null;
        }


        if (
            serialized ===
                this.lastRowSortStateFromDataView
        ) {

            return;
        }


        this.lastRowSortStateFromDataView =
            serialized;


        this.rowFieldSortState.clear();


        this.manualRowOrder.clear();


        if (
            serialized.length ===
                0
        ) {

            return;
        }


        try {


            const parsed =
                JSON.parse(
                    serialized
                ) as {
                    modes?: Record<
                        string,
                        unknown
                    >;
                    manual?: Record<
                        string,
                        Record<
                            string,
                            unknown
                        >
                    >;
                };


            Object.entries(
                parsed.modes ??
                {}
            ).forEach(
                (
                    [
                        fieldId,
                        mode
                    ]
                ) => {


                    if (
                        mode ===
                            "ascending" ||
                        mode ===
                            "descending" ||
                        mode ===
                            "manual"
                    ) {

                        this.rowFieldSortState.set(
                            fieldId,
                            mode
                        );
                    }
                }
            );


            Object.entries(
                parsed.manual ??
                {}
            ).forEach(
                (
                    [
                        fieldId,
                        parentMap
                    ]
                ) => {


                    const groups =
                        new Map<
                            string,
                            string[]
                        >();


                    Object.entries(
                        parentMap
                    ).forEach(
                        (
                            [
                                parentKey,
                                orderValue
                            ]
                        ) => {


                            if (
                                !Array.isArray(
                                    orderValue
                                )
                            ) {

                                return;
                            }


                            const seen =
                                new Set<string>();


                            const order =
                                orderValue
                                    .map(
                                        value =>
                                            String(
                                                value
                                            )
                                    )
                                    .filter(
                                        value => {


                                            if (
                                                value.length ===
                                                    0 ||
                                                seen.has(
                                                    value
                                                )
                                            ) {

                                                return false;
                                            }


                                            seen.add(
                                                value
                                            );

                                            return true;
                                        }
                                    );


                            groups.set(
                                parentKey,
                                order
                            );
                        }
                    );


                    this.manualRowOrder.set(
                        fieldId,
                        groups
                    );
                }
            );
        }
        catch {


            this.rowFieldSortState.clear();


            this.manualRowOrder.clear();
        }
    }


    private persistRowSortState(): void {


        const modes:
            Record<
                string,
                "ascending" |
                "descending" |
                "manual"
            > = {};


        this.rowFieldSortState
            .forEach(
                (
                    mode,
                    fieldId
                ) => {


                    modes[
                        fieldId
                    ] =
                        mode;
                }
            );


        const manual:
            Record<
                string,
                Record<
                    string,
                    string[]
                >
            > = {};


        this.manualRowOrder
            .forEach(
                (
                    parentMap,
                    fieldId
                ) => {


                    const groups:
                        Record<
                            string,
                            string[]
                        > = {};


                    parentMap
                        .forEach(
                            (
                                order,
                                parentKey
                            ) => {


                                groups[
                                    parentKey
                                ] =
                                    [
                                        ...order
                                    ];
                            }
                        );


                    manual[
                        fieldId
                    ] =
                        groups;
                }
            );


        const serialized =
            JSON.stringify({
                modes,
                manual
            });


        this.pendingRowSortState =
            serialized;


        this.host.persistProperties({

            merge: [
                {
                    objectName:
                        "internalState",

                    selector:
                        null,

                    properties: {
                        rowSortState:
                            serialized
                    }
                }
            ]
        });
    }


    /*
     * Hierarchy-aware Rows sorting.
     *
     * A lower-level field is always sorted only inside its parent group.
     * Manual order uses one independent sequence per parent group.
     */
    private sortRowsByFieldState(
        rows:
            MatrixAxisItem[],
        rowFieldIds:
            string[]
    ): MatrixAxisItem[] {


        const sortLevel =
            (
                items:
                    MatrixAxisItem[],
                level:
                    number
            ): MatrixAxisItem[] => {


                if (
                    level >=
                        rowFieldIds.length ||
                    items.length <=
                        1
                ) {

                    return [
                        ...items
                    ];
                }


                const groups =
                    new Map<
                        string,
                        {
                            valueKey:
                                string;
                            sortValue:
                                unknown;
                            items:
                                MatrixAxisItem[];
                        }
                    >();


                items.forEach(
                    item => {


                        const valueKey =
                            this.getAxisValueKey(
                                item,
                                level
                            );


                        const sortValue =
                            (
                                item.sortValues ??
                                item.values
                            )[
                                level
                            ];


                        let group =
                            groups.get(
                                valueKey
                            );


                        if (
                            !group
                        ) {

                            group = {
                                valueKey,
                                sortValue,
                                items:
                                    []
                            };


                            groups.set(
                                valueKey,
                                group
                            );
                        }


                        group.items.push(
                            item
                        );
                    }
                );


                const orderedGroups =
                    Array.from(
                        groups.values()
                    );


                const fieldId =
                    rowFieldIds[
                        level
                    ];


                const mode =
                    this.rowFieldSortState.get(
                        fieldId
                    ) ??
                    "default";


                if (
                    mode ===
                        "ascending" ||
                    mode ===
                        "descending"
                ) {


                    const direction =
                        mode ===
                            "descending"
                            ? -1
                            : 1;


                    orderedGroups.sort(
                        (
                            left,
                            right
                        ) =>
                            direction *
                            this.compareAxisValue(
                                left.sortValue,
                                right.sortValue
                            )
                    );
                }
                else if (
                    mode ===
                        "manual"
                ) {


                    const parentKey =
                        this.getAxisParentKey(
                            items[
                                0
                            ],
                            level
                        );


                    const manualOrder =
                        this.manualRowOrder
                            .get(
                                fieldId
                            )
                            ?.get(
                                parentKey
                            ) ??
                        [];


                    const manualPosition =
                        new Map<
                            string,
                            number
                        >();


                    manualOrder.forEach(
                        (
                            valueKey,
                            index
                        ) => {


                            manualPosition.set(
                                valueKey,
                                index
                            );
                        }
                    );


                    const originalPosition =
                        new Map<
                            string,
                            number
                        >();


                    orderedGroups.forEach(
                        (
                            group,
                            index
                        ) => {


                            originalPosition.set(
                                group.valueKey,
                                index
                            );
                        }
                    );


                    orderedGroups.sort(
                        (
                            left,
                            right
                        ) => {


                            const leftPosition =
                                manualPosition.get(
                                    left.valueKey
                                );


                            const rightPosition =
                                manualPosition.get(
                                    right.valueKey
                                );


                            if (
                                leftPosition !==
                                    undefined &&
                                rightPosition !==
                                    undefined
                            ) {

                                return (
                                    leftPosition -
                                    rightPosition
                                );
                            }


                            if (
                                leftPosition !==
                                    undefined
                            ) {

                                return -1;
                            }


                            if (
                                rightPosition !==
                                    undefined
                            ) {

                                return 1;
                            }


                            return (
                                (
                                    originalPosition.get(
                                        left.valueKey
                                    ) ??
                                    0
                                ) -
                                (
                                    originalPosition.get(
                                        right.valueKey
                                    ) ??
                                    0
                                )
                            );
                        }
                    );
                }


                const result:
                    MatrixAxisItem[] =
                    [];


                orderedGroups.forEach(
                    group => {


                        result.push(
                            ...sortLevel(
                                group.items,
                                level +
                                    1
                            )
                        );
                    }
                );


                return result;
            };


        return sortLevel(
            rows,
            0
        );
    }


    private getAxisValueKey(
        item:
            MatrixAxisItem,
        level:
            number
    ): string {


        const rawValue =
            (
                item.sortValues ??
                item.values
            )[
                level
            ];


        return JSON.stringify([
            typeof rawValue,
            rawValue,
            item.values[
                level
            ] ??
            ""
        ]);
    }


    private getAxisParentKey(
        item:
            MatrixAxisItem,
        level:
            number
    ): string {


        const rawValues =
            item.sortValues ??
            item.values;


        const parts:
            unknown[] =
            [];


        for (
            let index =
                0;
            index <
                level;
            index +=
                1
        ) {


            parts.push([
                typeof rawValues[
                    index
                ],
                rawValues[
                    index
                ],
                item.values[
                    index
                ] ??
                ""
            ]);
        }


        return JSON.stringify(
            parts
        );
    }


    /*
     * Merge the newly dragged visible order with a possibly larger
     * persisted order so values hidden by filters keep their place.
     */
    private setManualRowOrder(
        fieldId:
            string,
        parentKey:
            string,
        visibleOrder:
            string[]
    ): void {


        let parentMap =
            this.manualRowOrder.get(
                fieldId
            );


        if (
            !parentMap
        ) {

            parentMap =
                new Map<
                    string,
                    string[]
                >();


            this.manualRowOrder.set(
                fieldId,
                parentMap
            );
        }


        const existing =
            parentMap.get(
                parentKey
            ) ??
            [];


        const visibleSet =
            new Set(
                visibleOrder
            );


        let merged:
            string[];


        if (
            existing.length ===
                0
        ) {

            merged =
                [
                    ...visibleOrder
                ];
        }
        else {


            const queue =
                [
                    ...visibleOrder
                ];


            merged =
                existing.map(
                    key => {


                        if (
                            !visibleSet.has(
                                key
                            )
                        ) {

                            return key;
                        }


                        return (
                            queue.shift() ??
                            key
                        );
                    }
                );


            merged.push(
                ...queue
            );
        }


        const seen =
            new Set<string>();


        parentMap.set(
            parentKey,
            merged.filter(
                key => {


                    if (
                        seen.has(
                            key
                        )
                    ) {

                        return false;
                    }


                    seen.add(
                        key
                    );

                    return true;
                }
            )
        );


        this.persistRowSortState();
    }


    /*
     * Restore the sort mode of each Columns field from PBIX metadata.
     */
    private restoreColumnFieldSortState(
        dataView:
            powerbi.DataView
    ): void {


        const rawValue =
            dataView
                .metadata
                .objects
                ?.internalState
                ?.columnSortState;


        const serialized =
            typeof rawValue ===
                "string"
                ? rawValue
                : "";


        if (
            this.pendingColumnFieldSortState !==
                null
        ) {


            if (
                serialized ===
                this.pendingColumnFieldSortState
            ) {

                this.lastColumnFieldSortStateFromDataView =
                    serialized;


                this.pendingColumnFieldSortState =
                    null;


                return;
            }


            if (
                serialized ===
                this.lastColumnFieldSortStateFromDataView
            ) {

                return;
            }


            this.pendingColumnFieldSortState =
                null;
        }


        if (
            serialized ===
                this.lastColumnFieldSortStateFromDataView
        ) {

            return;
        }


        this.lastColumnFieldSortStateFromDataView =
            serialized;


        this.columnFieldSortState.clear();


        if (
            serialized.length ===
                0
        ) {

            return;
        }


        try {


            const parsed =
                JSON.parse(
                    serialized
                ) as Record<string, unknown>;


            Object.entries(
                parsed
            ).forEach(
                (
                    [
                        fieldId,
                        mode
                    ]
                ) => {


                    if (
                        mode ===
                            "ascending" ||
                        mode ===
                            "descending"
                    ) {

                        this.columnFieldSortState.set(
                            fieldId,
                            mode
                        );
                    }
                }
            );
        }
        catch {


            this.columnFieldSortState.clear();
        }
    }


    private persistColumnFieldSortState(): void {


        const state:
            Record<
                string,
                "ascending" | "descending"
            > = {};


        this.columnFieldSortState
            .forEach(
                (
                    mode,
                    fieldId
                ) => {


                    state[
                        fieldId
                    ] =
                        mode;
                }
            );


        const serialized =
            JSON.stringify(
                state
            );


        this.pendingColumnFieldSortState =
            serialized;


        this.host.persistProperties({

            merge: [
                {
                    objectName:
                        "internalState",

                    selector:
                        null,

                    properties: {
                        columnSortState:
                            serialized
                    }
                }
            ]
        });
    }


    /*
     * Sort each hierarchy level only inside its parent group.
     *
     * Example:
     * Patient ID sorted Ascending moves whole Patient ID groups.
     * Internal Batch ID sorted Ascending only sorts batches inside
     * each Patient ID and never mixes batches between patients.
     */
    private sortColumnsByFieldState(
        columns:
            MatrixAxisItem[],
        columnFieldIds:
            string[]
    ): MatrixAxisItem[] {


        const sortLevel =
            (
                items:
                    MatrixAxisItem[],
                level:
                    number
            ): MatrixAxisItem[] => {


                if (
                    level >=
                        columnFieldIds.length ||
                    items.length <=
                        1
                ) {

                    return [
                        ...items
                    ];
                }


                const groups =
                    new Map<
                        string,
                        {
                            sortValue:
                                unknown;
                            items:
                                MatrixAxisItem[];
                        }
                    >();


                items.forEach(
                    item => {


                        const sortValue =
                            (
                                item.sortValues ??
                                item.values
                            )[
                                level
                            ];


                        const groupKey =
                            JSON.stringify([
                                typeof sortValue,
                                sortValue,
                                item.values[
                                    level
                                ] ?? ""
                            ]);


                        let group =
                            groups.get(
                                groupKey
                            );


                        if (
                            !group
                        ) {

                            group = {
                                sortValue,
                                items:
                                    []
                            };


                            groups.set(
                                groupKey,
                                group
                            );
                        }


                        group.items.push(
                            item
                        );
                    }
                );


                const orderedGroups =
                    Array.from(
                        groups.values()
                    );


                const mode =
                    this.columnFieldSortState.get(
                        columnFieldIds[
                            level
                        ]
                    ) ??
                    "default";


                if (
                    mode ===
                        "ascending" ||
                    mode ===
                        "descending"
                ) {


                    const direction =
                        mode ===
                            "descending"
                            ? -1
                            : 1;


                    orderedGroups.sort(
                        (
                            left,
                            right
                        ) =>
                            direction *
                            this.compareAxisValue(
                                left.sortValue,
                                right.sortValue
                            )
                    );
                }


                const result:
                    MatrixAxisItem[] =
                    [];


                orderedGroups.forEach(
                    group => {


                        result.push(
                            ...sortLevel(
                                group.items,
                                level +
                                    1
                            )
                        );
                    }
                );


                return result;
            };


        return sortLevel(
            columns,
            0
        );
    }


    private getAxisSortValue(
        value:
            unknown,
        column:
            powerbi.DataViewMetadataColumn
    ): unknown {


        if (
            value ===
                null ||
            value ===
                undefined
        ) {

            return null;
        }


        if (
            column.type &&
            column.type.dateTime
        ) {


            const dateValue =
                value instanceof Date
                    ? value
                    : new Date(
                        String(
                            value
                        )
                    );


            if (
                !Number.isNaN(
                    dateValue.getTime()
                )
            ) {

                return dateValue.getTime();
            }
        }


        return value;
    }


    private compareAxisValue(
        left:
            unknown,
        right:
            unknown
    ): number {


        const leftEmpty =
            left ===
                null ||
            left ===
                undefined ||
            String(
                left
            ).length ===
                0;


        const rightEmpty =
            right ===
                null ||
            right ===
                undefined ||
            String(
                right
            ).length ===
                0;


        if (
            leftEmpty &&
            rightEmpty
        ) {

            return 0;
        }


        if (
            leftEmpty
        ) {

            return 1;
        }


        if (
            rightEmpty
        ) {

            return -1;
        }


        if (
            typeof left ===
                "number" &&
            typeof right ===
                "number"
        ) {

            return (
                left -
                right
            );
        }


        if (
            typeof left ===
                "boolean" &&
            typeof right ===
                "boolean"
        ) {

            return Number(
                left
            ) -
            Number(
                right
            );
        }


        return String(
            left
        ).localeCompare(
            String(
                right
            ),
            this.locale,
            {
                numeric:
                    true,

                sensitivity:
                    "base"
            }
        );
    }


    /*
     * Load manually resized widths from Power BI metadata.
     *
     * Only manual widths are persisted. Auto widths are always
     * recalculated from the current content.
     */
    private restoreLayoutState(
        dataView: powerbi.DataView
    ): void {


        const rawValue =
            dataView
                .metadata
                .objects
                ?.internalState
                ?.layoutState;


        const serialized =
            typeof rawValue ===
                "string"
                ? rawValue
                : "";


        /*
         * persistProperties can cause another update. If Power BI
         * briefly sends the previous state while our new state is
         * pending, keep the local widths instead of reverting them.
         */
        if (
            this.pendingLayoutState !==
                null
        ) {


            if (
                serialized ===
                this.pendingLayoutState
            ) {

                this.lastLayoutStateFromDataView =
                    serialized;


                this.pendingLayoutState =
                    null;


                return;
            }


            if (
                serialized ===
                this.lastLayoutStateFromDataView
            ) {

                return;
            }


            this.pendingLayoutState =
                null;
        }


        if (
            serialized ===
            this.lastLayoutStateFromDataView
        ) {

            return;
        }


        this.lastLayoutStateFromDataView =
            serialized;


        this.rowColumnWidths =
            [];


        this.dataColumnWidths.clear();


        this.manuallyResizedRowColumns.clear();


        this.manuallyResizedDataColumns.clear();


        if (
            serialized.length ===
                0
        ) {

            return;
        }


        try {


            const parsed =
                JSON.parse(
                    serialized
                ) as Partial<PersistedLayoutState>;


            if (
                parsed.version !==
                    1
            ) {

                return;
            }


            if (
                Array.isArray(
                    parsed.rowWidths
                )
            ) {


                parsed.rowWidths.forEach(
                    entry => {


                        if (
                            !Array.isArray(
                                entry
                            ) ||
                            entry.length !==
                                2
                        ) {

                            return;
                        }


                        const level =
                            Number(
                                entry[0]
                            );


                        const width =
                            this.normalizePersistedWidth(
                                entry[1]
                            );


                        if (
                            !Number.isInteger(
                                level
                            ) ||
                            level <
                                0 ||
                            width ===
                                null
                        ) {

                            return;
                        }


                        this.rowColumnWidths[
                            level
                        ] =
                            width;


                        this.manuallyResizedRowColumns.add(
                            level
                        );
                    }
                );
            }


            if (
                Array.isArray(
                    parsed.dataWidths
                )
            ) {


                parsed.dataWidths.forEach(
                    entry => {


                        if (
                            !Array.isArray(
                                entry
                            ) ||
                            entry.length !==
                                2
                        ) {

                            return;
                        }


                        const key =
                            String(
                                entry[0]
                            );


                        const width =
                            this.normalizePersistedWidth(
                                entry[1]
                            );


                        if (
                            key.length ===
                                0 ||
                            width ===
                                null
                        ) {

                            return;
                        }


                        this.dataColumnWidths.set(
                            key,
                            width
                        );


                        this.manuallyResizedDataColumns.add(
                            key
                        );
                    }
                );
            }
        }
        catch {


            /*
             * Ignore malformed persisted state and fall back to Auto width.
             */
            this.rowColumnWidths =
                [];


            this.dataColumnWidths.clear();


            this.manuallyResizedRowColumns.clear();


            this.manuallyResizedDataColumns.clear();
        }
    }


    /*
     * Persist only user-resized widths.
     *
     * Auto width remains data-driven and therefore does not need to be
     * stored in the PBIX.
     */
    private persistLayoutState(): void {


        const rowWidths:
            Array<[number, number]> =
            [];


        this.manuallyResizedRowColumns
            .forEach(
                level => {


                    const width =
                        this.rowColumnWidths[
                            level
                        ];


                    if (
                        Number.isFinite(
                            width
                        )
                    ) {

                        rowWidths.push([
                            level,
                            width
                        ]);
                    }
                }
            );


        rowWidths.sort(
            (
                left,
                right
            ) =>
                left[0] -
                right[0]
        );


        const dataWidths:
            Array<[string, number]> =
            [];


        this.manuallyResizedDataColumns
            .forEach(
                key => {


                    const width =
                        this.dataColumnWidths.get(
                            key
                        );


                    if (
                        width !==
                            undefined &&
                        Number.isFinite(
                            width
                        )
                    ) {

                        dataWidths.push([
                            key,
                            width
                        ]);
                    }
                }
            );


        dataWidths.sort(
            (
                left,
                right
            ) =>
                left[0].localeCompare(
                    right[0]
                )
        );


        const state:
            PersistedLayoutState = {

                version:
                    1,

                rowWidths,

                dataWidths
            };


        const serialized =
            JSON.stringify(
                state
            );


        this.pendingLayoutState =
            serialized;


        this.host.persistProperties({

            merge: [
                {
                    objectName:
                        "internalState",

                    selector:
                        null,

                    properties: {
                        layoutState:
                            serialized
                    }
                }
            ]
        });
    }


    /*
     * Persisted metadata is user-editable at the file level, so validate
     * widths before applying them.
     */
    private normalizePersistedWidth(
        value: unknown
    ): number | null {


        const width =
            Number(
                value
            );


        if (
            !Number.isFinite(
                width
            )
        ) {

            return null;
        }


        return Math.min(
            600,
            Math.max(
                55,
                Math.round(
                    width
                )
            )
        );
    }


    private addColumnResizeHandle(
        headerCell: HTMLTableCellElement,
        getCurrentWidth: () => number,
        setWidth: (width: number, isManual: boolean) => void,
        defaultWidth: number,
        minimumWidth: number,
        maximumWidth: number
    ): void {


        const handle =
            document.createElement(
                "div"
            );


        handle.setAttribute(
            "aria-label",
            "Resize column"
        );


        handle.title =
            "Drag to resize column. Double-click to reset.";


        handle.style.position =
            "absolute";


        handle.style.top =
            "0";


        handle.style.right =
            "-3px";


        handle.style.width =
            "7px";


        handle.style.height =
            "100%";


        handle.style.cursor =
            "col-resize";


        handle.style.zIndex =
            "30";


        handle.style.touchAction =
            "none";


        handle.style.background =
            "transparent";


        const guide =
            document.createElement(
                "div"
            );


        guide.style.position =
            "absolute";


        guide.style.top =
            "0";


        guide.style.bottom =
            "0";


        guide.style.left =
            "3px";


        guide.style.width =
            "1px";


        guide.style.background =
            "transparent";


        guide.style.pointerEvents =
            "none";


        handle.appendChild(
            guide
        );


        handle.addEventListener(
            "pointerenter",
            () => {

                guide.style.background =
                    this.isHighContrast
                        ? this.highContrastForeground
                        : "#8A8A8A";
            }
        );


        handle.addEventListener(
            "pointerleave",
            () => {

                guide.style.background =
                    "transparent";
            }
        );


        handle.addEventListener(
            "pointerdown",
            (
                event:
                    PointerEvent
            ) => {


                event.preventDefault();
                event.stopPropagation();


                const startX =
                    event.clientX;


                const startWidth =
                    getCurrentWidth();


                const previousUserSelect =
                    this.target.style.userSelect;


                this.target.style.userSelect =
                    "none";


                guide.style.background =
                    this.isHighContrast
                        ? this.highContrastForeground
                        : "#666666";


                handle.setPointerCapture(
                    event.pointerId
                );


                const onPointerMove =
                    (
                        moveEvent:
                            PointerEvent
                    ): void => {


                        const delta =
                            moveEvent.clientX -
                            startX;


                        const nextWidth =
                            Math.min(
                                maximumWidth,
                                Math.max(
                                    minimumWidth,
                                    Math.round(
                                        startWidth +
                                        delta
                                    )
                                )
                            );


                        setWidth(
                            nextWidth,
                            true
                        );
                    };


                const finishResize =
                    (
                        finishEvent:
                            PointerEvent
                    ): void => {


                        handle.removeEventListener(
                            "pointermove",
                            onPointerMove
                        );


                        handle.removeEventListener(
                            "pointerup",
                            finishResize
                        );


                        handle.removeEventListener(
                            "pointercancel",
                            finishResize
                        );


                        if (
                            handle.hasPointerCapture(
                                finishEvent.pointerId
                            )
                        ) {

                            handle.releasePointerCapture(
                                finishEvent.pointerId
                            );
                        }


                        this.target.style.userSelect =
                            previousUserSelect;


                        guide.style.background =
                            "transparent";


                        this.persistLayoutState();
                    };


                handle.addEventListener(
                    "pointermove",
                    onPointerMove
                );


                handle.addEventListener(
                    "pointerup",
                    finishResize
                );


                handle.addEventListener(
                    "pointercancel",
                    finishResize
                );
            }
        );


        handle.addEventListener(
            "dblclick",
            event => {


                event.preventDefault();
                event.stopPropagation();


                setWidth(
                    defaultWidth,
                    false
                );


                this.persistLayoutState();
            }
        );


        headerCell.appendChild(
            handle
        );
    }


    private createMark(
        mark: MarkPoint
    ): HTMLElement {


        const element =
            document.createElement(
                "div"
            );


        const appearance =
            this.formattingSettings
                .matrixAppearance;


        const size =
            appearance
                .markSize
                .value;


        const shape =
            this.getStatusShape(
                mark.status
            );


        element.style.width =
            `${size}px`;


        element.style.height =
            `${size}px`;


        element.style.minWidth =
            `${size}px`;


        element.style.minHeight =
            `${size}px`;


        element.style.backgroundColor =
            this.isHighContrast
                ? this.highContrastBackground
                : this.getStatusColor(
                    mark.status
                );


        if (
            this.isHighContrast
        ) {

            element.style.border =
                `2px solid ${this.highContrastForeground}`;
        }


        element.style.cursor =
            this.allowInteractions
                ? "pointer"
                : "default";


        element.style.boxSizing =
            "border-box";


        /*
         * Keyboard and screen-reader accessibility.
         *
         * div elements are not tabbable by default, so every mark becomes
         * a real keyboard focus target. The aria-label is built from the
         * same information that is available in the tooltip.
         */
        element.tabIndex =
            0;


        element.setAttribute(
            "role",
            "button"
        );


        element.setAttribute(
            "aria-label",
            this.getMarkAriaLabel(
                mark
            )
        );


        switch (
            shape
        ) {


            case "circle":

                element.style.borderRadius =
                    "50%";

                break;


            case "diamond":

                element.style.clipPath =
                    "polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%)";

                break;


            case "triangle":

                element.style.clipPath =
                    "polygon(50% 0%, 100% 100%, 0% 100%)";

                break;


            default:

                element.style.borderRadius =
                    "0";

                break;
        }


        /*
         * Standard tooltip + Report Page Tooltip context.
         */
        element.addEventListener(
            "pointerover",
            (
                event:
                    PointerEvent
            ) => {


                if (
                    !this.host
                        .tooltipService
                        .enabled()
                ) {

                    return;
                }


                this.host
                    .tooltipService
                    .show({

                        coordinates: [
                            event.clientX,
                            event.clientY
                        ],

                        isTouchEvent:
                            event.pointerType ===
                            "touch",

                        dataItems:
                            mark.tooltipItems,

                        identities: [
                            mark.selectionId
                        ]
                    });
            }
        );


        element.addEventListener(
            "pointermove",
            (
                event:
                    PointerEvent
            ) => {


                if (
                    !this.host
                        .tooltipService
                        .enabled()
                ) {

                    return;
                }


                this.host
                    .tooltipService
                    .move({

                        coordinates: [
                            event.clientX,
                            event.clientY
                        ],

                        isTouchEvent:
                            event.pointerType ===
                            "touch",

                        dataItems:
                            mark.tooltipItems,

                        identities: [
                            mark.selectionId
                        ]
                    });
            }
        );


        element.addEventListener(
            "pointerout",
            (
                event:
                    PointerEvent
            ) => {


                this.host
                    .tooltipService
                    .hide({

                        isTouchEvent:
                            event.pointerType ===
                            "touch",

                        immediately:
                            false
                    });
            }
        );


        /*
         * Selection / cross-filtering.
         */
        element.addEventListener(
            "click",
            (
                event:
                    MouseEvent
            ) => {


                event.stopPropagation();


                if (
                    !this.allowInteractions
                ) {

                    return;
                }


                this.selectionManager
                    .select(
                        mark.selectionId,
                        event.ctrlKey
                    );
            }
        );


        /*
         * Native Power BI context menu for a specific mark.
         */
        element.addEventListener(
            "contextmenu",
            (
                event:
                    MouseEvent
            ) => {


                if (
                    !this.allowInteractions
                ) {

                    return;
                }


                event.preventDefault();
                event.stopPropagation();


                this.selectionManager
                    .showContextMenu(
                        mark.selectionId,
                        {
                            x:
                                event.clientX,

                            y:
                                event.clientY
                        }
                    );
            }
        );


        /*
         * Enter / Space perform the same action as mouse click.
         * Ctrl keeps the existing multi-select behavior.
         *
         * Shift+F10 and the ContextMenu key open the native
         * Power BI context menu for keyboard users.
         */
        element.addEventListener(
            "keydown",
            (
                event:
                    KeyboardEvent
            ) => {


                const isContextMenuKey =
                    event.key ===
                        "ContextMenu" ||
                    (
                        event.shiftKey &&
                        event.key ===
                            "F10"
                    );


                if (
                    isContextMenuKey
                ) {


                    event.preventDefault();
                    event.stopPropagation();


                    if (
                        !this.allowInteractions
                    ) {

                        return;
                    }


                    const rect =
                        element
                            .getBoundingClientRect();


                    this.selectionManager
                        .showContextMenu(
                            mark.selectionId,
                            {
                                x:
                                    rect.left +
                                    (
                                        rect.width /
                                        2
                                    ),

                                y:
                                    rect.top +
                                    (
                                        rect.height /
                                        2
                                    )
                            }
                        );


                    return;
                }


                if (
                    event.key !==
                        "Enter" &&
                    event.key !==
                        " "
                ) {

                    return;
                }


                event.preventDefault();
                event.stopPropagation();


                if (
                    !this.allowInteractions
                ) {

                    return;
                }


                this.selectionManager
                    .select(
                        mark.selectionId,
                        event.ctrlKey ||
                        event.metaKey
                    );
            }
        );


        /*
         * Keep a clearly visible focus indicator.
         * High-contrast-aware styling will be added in the next
         * accessibility step.
         */
        element.addEventListener(
            "focus",
            () => {


                element.style.outline =
                    `2px solid ${
                        this.isHighContrast
                            ? this.highContrastSelected
                            : "#000000"
                    }`;


                element.style.outlineOffset =
                    "2px";
            }
        );


        element.addEventListener(
            "blur",
            () => {


                element.style.outline =
                    "none";


                element.style.outlineOffset =
                    "0";
            }
        );


        return element;
    }


    /*
     * Screen-reader text for a mark.
     *
     * Tooltip items already contain Status, Mark Order and any fields
     * assigned to the Tooltips role, so reusing them keeps accessible
     * text aligned with what sighted users see in the tooltip.
     */
    private getMarkAriaLabel(
        mark: MarkPoint
    ): string {


        const parts =
            mark.tooltipItems
                .filter(
                    item =>
                        item.value !==
                            null &&
                        item.value !==
                            undefined &&
                        String(
                            item.value
                        ).trim() !==
                            ""
                )
                .map(
                    item =>
                        `${item.displayName}: ${String(
                            item.value
                        )}`
                );


        if (
            parts.length === 0
        ) {

            parts.push(
                `Status: ${mark.status}`
            );
        }


        return parts.join(
            ", "
        );
    }


    /*
     * Format values according to Power BI metadata.
     *
     * Date/DateTime values can arrive as ISO timestamp strings,
     * so convert them back to Date before applying the format.
     */
    private formatValue(
        value: powerbi.PrimitiveValue,
        column: powerbi.DataViewMetadataColumn
    ): string {


        if (
            value === null ||
            value === undefined
        ) {

            return "";
        }


        if (
            column.type &&
            column.type.dateTime
        ) {


            let dateValue: Date;


            if (
                value instanceof Date
            ) {

                dateValue =
                    value;
            }
            else {

                dateValue =
                    new Date(
                        String(value)
                    );
            }


            if (
                !Number.isNaN(
                    dateValue.getTime()
                )
            ) {


                const format =
                    column.format ||
                    "dd.MM.yyyy";


                return valueFormatter.format(
                    dateValue,
                    format,
                    false,
                    this.locale
                );
            }
        }


        return valueFormatter.format(
            value,
            column.format,
            false,
            this.locale
        );
    }


    private rawValueToText(
        value:
            powerbi.PrimitiveValue
    ): string {


        if (
            value === null ||
            value === undefined
        ) {

            return "";
        }


        return String(
            value
        );
    }


    /*
     * Returns the shape configured for a specific status.
     *
     * Unknown statuses fall back to Matrix appearance ->
     * Default mark shape.
     */
    private getStatusShape(
        status: string
    ): string {


        const shapes =
            this.formattingSettings
                .statusShapes;


        switch (
            status
        ) {


            case "Analyse durchgeführt":

                return String(
                    shapes
                        .analyseDurchgefuehrtShape
                        .value
                        .value
                );


            case "Analyse freigegeben":

                return String(
                    shapes
                        .analyseFreigegebenShape
                        .value
                        .value
                );


            case "Analyse gereviewed":

                return String(
                    shapes
                        .analyseGereviewedShape
                        .value
                        .value
                );


            case "Analyse r/c + Batch closed":

                return String(
                    shapes
                        .analyseRcBatchClosedShape
                        .value
                        .value
                );


            case "Durchzuführende Analyse":

                return String(
                    shapes
                        .durchzufuehrendeAnalyseShape
                        .value
                        .value
                );


            case "Investigation open":

                return String(
                    shapes
                        .investigationOpenShape
                        .value
                        .value
                );


            case "Sample rejected/canceled":

                return String(
                    shapes
                        .sampleRejectedCanceledShape
                        .value
                        .value
                );


            case "Analyse in Bearbeitung":

                return String(
                    shapes
                        .analyseInBearbeitungShape
                        .value
                        .value
                );


            case "Analyse rejected/canceled":

                return String(
                    shapes
                        .analyseRejectedCanceledShape
                        .value
                        .value
                );


            case "Not mapped":

                return String(
                    shapes
                        .notMappedShape
                        .value
                        .value
                );


            default:

                return String(
                    this.formattingSettings
                        .matrixAppearance
                        .markShape
                        .value
                        .value
                );
        }
    }


    private getStatusColor(
        status: string
    ): string {


        const colors =
            this.formattingSettings
                .statusColors;


        switch (
            status
        ) {


            case "Analyse durchgeführt":

                return colors
                    .analyseDurchgefuehrt
                    .value.value;


            case "Analyse freigegeben":

                return colors
                    .analyseFreigegeben
                    .value.value;


            case "Analyse gereviewed":

                return colors
                    .analyseGereviewed
                    .value.value;


            case "Analyse r/c + Batch closed":

                return colors
                    .analyseRcBatchClosed
                    .value.value;


            case "Durchzuführende Analyse":

                return colors
                    .durchzufuehrendeAnalyse
                    .value.value;


            case "Investigation open":

                return colors
                    .investigationOpen
                    .value.value;


            case "Sample rejected/canceled":

                return colors
                    .sampleRejectedCanceled
                    .value.value;


            case "Analyse in Bearbeitung":

                return colors
                    .analyseInBearbeitung
                    .value.value;


            case "Analyse rejected/canceled":

                return colors
                    .analyseRejectedCanceled
                    .value.value;


            case "Not mapped":

                return colors
                    .notMapped
                    .value.value;


            default:

                return "#808080";
        }
    }


    private shouldRenderRowCell(
        rows:
            MatrixAxisItem[],
        rowIndex:
            number,
        level:
            number
    ): boolean {


        if (
            rowIndex === 0
        ) {

            return true;
        }


        const current =
            rows[
                rowIndex
            ].values;


        const previous =
            rows[
                rowIndex - 1
            ].values;


        for (
            let i = 0;
            i <= level;
            i++
        ) {


            if (
                current[i] !==
                previous[i]
            ) {

                return true;
            }
        }


        return false;
    }


    private getRowSpan(
        rows:
            MatrixAxisItem[],
        startIndex:
            number,
        level:
            number
    ): number {


        let span =
            1;


        const current =
            rows[
                startIndex
            ].values;


        for (
            let i =
                startIndex + 1;
            i <
            rows.length;
            i++
        ) {


            const candidate =
                rows[
                    i
                ].values;


            let sameGroup =
                true;


            for (
                let hierarchyLevel =
                    0;
                hierarchyLevel <=
                level;
                hierarchyLevel++
            ) {


                if (
                    candidate[
                        hierarchyLevel
                    ] !==
                    current[
                        hierarchyLevel
                    ]
                ) {

                    sameGroup =
                        false;

                    break;
                }
            }


            if (
                !sameGroup
            ) {

                break;
            }


            span++;
        }


        return span;
    }


    private getColumnSpan(
        columns:
            MatrixAxisItem[],
        startIndex:
            number,
        level:
            number
    ): number {


        let span =
            1;


        const current =
            columns[
                startIndex
            ].values;


        for (
            let i =
                startIndex + 1;
            i <
            columns.length;
            i++
        ) {


            const candidate =
                columns[
                    i
                ].values;


            let sameGroup =
                true;


            for (
                let hierarchyLevel =
                    0;
                hierarchyLevel <=
                level;
                hierarchyLevel++
            ) {


                if (
                    candidate[
                        hierarchyLevel
                    ] !==
                    current[
                        hierarchyLevel
                    ]
                ) {

                    sameGroup =
                        false;

                    break;
                }
            }


            if (
                !sameGroup
            ) {

                break;
            }


            span++;
        }


        return span;
    }


    private getRoleIndexes(
        columns:
            powerbi.DataViewMetadataColumn[],
        roleName:
            string
    ): number[] {


        const roleColumns =
            columns
                .map(
                    (
                        column,
                        index
                    ) => {


                        const runtimeColumn =
                            column as
                                RuntimeRoleIndexColumn;


                        const rolePosition =
                            runtimeColumn
                                .rolesIndex?.[
                                    roleName
                                ]?.[0];


                        return {
                            index,
                            matches:
                                Boolean(
                                    column.roles &&
                                    column.roles[
                                        roleName
                                    ]
                                ),
                            rolePosition:
                                typeof rolePosition ===
                                    "number"
                                    ? rolePosition
                                    : null
                        };
                    }
                )
                .filter(
                    item =>
                        item.matches
                );


        /*
         * When Power BI provides rolesIndex, it reflects the current
         * order in the field well. This fixes the case where a field is
         * inserted or dragged into the middle but metadata.columns keeps
         * the older insertion order.
         *
         * If rolesIndex is unavailable, preserve the existing safe
         * behavior and use metadata order.
         */
        roleColumns.sort(
            (
                left,
                right
            ) => {


                if (
                    left.rolePosition !==
                        null &&
                    right.rolePosition !==
                        null
                ) {

                    return (
                        left.rolePosition -
                        right.rolePosition
                    );
                }


                if (
                    left.rolePosition !==
                        null
                ) {

                    return -1;
                }


                if (
                    right.rolePosition !==
                        null
                ) {

                    return 1;
                }


                return (
                    left.index -
                    right.index
                );
            }
        );


        return roleColumns.map(
            item =>
                item.index
        );
    }



    private getRoleIndex(
        columns:
            powerbi.DataViewMetadataColumn[],
        roleName:
            string
    ): number {


        return columns
            .findIndex(
                column =>
                    Boolean(
                        column.roles &&
                        column.roles[
                            roleName
                        ]
                    )
            );
    }


    private compareOrder(
        a:
            string,
        b:
            string
    ): number {


        const numberA =
            Number(a);


        const numberB =
            Number(b);


        if (
            !Number.isNaN(
                numberA
            ) &&
            !Number.isNaN(
                numberB
            )
        ) {

            return (
                numberA -
                numberB
            );
        }


        return a.localeCompare(
            b
        );
    }


    private styleHeader(
        element:
            HTMLTableCellElement,
        background:
            string,
        textColor:
            string,
        showGrid:
            boolean,
        gridColor:
            string,
        fontSize:
            number
    ): void {


        element.style.padding =
            "6px 8px";


        element.style.fontFamily =
            "Segoe UI, Arial, sans-serif";


        element.style.fontSize =
            `${fontSize}px`;


        element.style.fontWeight =
            "600";


        element.style.background =
            background;


        element.style.color =
            textColor;


        if (
            showGrid
        ) {

            element.style.borderBottom =
                `1px solid ${gridColor}`;


            element.style.borderRight =
                `1px solid ${gridColor}`;
        }
    }


    private showMessage(
        message:
            string
    ): void {


        /*
         * Landing / empty state.
         *
         * Built with safe DOM methods and designed to remain readable
         * in Power BI high-contrast mode.
         */
        const container =
            document.createElement(
                "div"
            );


        container.setAttribute(
            "role",
            "status"
        );


        container.setAttribute(
            "aria-live",
            "polite"
        );


        container.style.width =
            "100%";


        container.style.height =
            "100%";


        container.style.boxSizing =
            "border-box";


        container.style.display =
            "flex";


        container.style.flexDirection =
            "column";


        container.style.alignItems =
            "center";


        container.style.justifyContent =
            "center";


        container.style.gap =
            "10px";


        container.style.padding =
            "24px";


        container.style.textAlign =
            "center";


        container.style.fontFamily =
            "Segoe UI, Arial, sans-serif";


        container.style.color =
            this.isHighContrast
                ? this.highContrastForeground
                : "#333333";


        container.style.background =
            this.isHighContrast
                ? this.highContrastBackground
                : "#FFFFFF";


        const icon =
            document.createElement(
                "div"
            );


        icon.setAttribute(
            "aria-hidden",
            "true"
        );


        icon.style.display =
            "grid";


        icon.style.gridTemplateColumns =
            "repeat(3, 12px)";


        icon.style.gridTemplateRows =
            "repeat(3, 12px)";


        icon.style.gap =
            "3px";


        icon.style.padding =
            "8px";


        icon.style.border =
            `1px solid ${
                this.isHighContrast
                    ? this.highContrastForeground
                    : "#C8C8C8"
            }`;


        icon.style.borderRadius =
            "6px";


        for (
            let index =
                0;
            index <
                9;
            index +=
                1
        ) {


            const cell =
                document.createElement(
                    "span"
                );


            cell.style.display =
                "block";


            cell.style.width =
                "12px";


            cell.style.height =
                "12px";


            cell.style.boxSizing =
                "border-box";


            cell.style.border =
                `1px solid ${
                    this.isHighContrast
                        ? this.highContrastForeground
                        : "#8A8A8A"
                }`;


            if (
                !this.isHighContrast &&
                (
                    index ===
                        1 ||
                    index ===
                        4 ||
                    index ===
                        8
                )
            ) {

                cell.style.background =
                    "#E8E8E8";
            }


            icon.appendChild(
                cell
            );
        }


        const title =
            document.createElement(
                "div"
            );


        title.textContent =
            "Multi Mark Matrix";


        title.style.fontSize =
            "16px";


        title.style.fontWeight =
            "600";


        const instruction =
            document.createElement(
                "div"
            );


        instruction.textContent =
            message;


        instruction.style.maxWidth =
            "420px";


        instruction.style.fontSize =
            "12px";


        instruction.style.lineHeight =
            "1.4";


        const optional =
            document.createElement(
                "div"
            );


        optional.textContent =
            "Optional: add Mark Order and Tooltips for richer mark behavior.";


        optional.style.maxWidth =
            "420px";


        optional.style.fontSize =
            "11px";


        optional.style.lineHeight =
            "1.35";


        optional.style.color =
            this.isHighContrast
                ? this.highContrastForeground
                : "#666666";


        container.appendChild(
            icon
        );


        container.appendChild(
            title
        );


        container.appendChild(
            instruction
        );


        container.appendChild(
            optional
        );


        this.target.appendChild(
            container
        );
    }

}
