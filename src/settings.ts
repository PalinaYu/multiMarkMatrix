"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard =
    formattingSettings.SimpleCard;

import FormattingSettingsSlice =
    formattingSettings.Slice;

import FormattingSettingsModel =
    formattingSettings.Model;


const shapeItems = [
    {
        displayName: "Square",
        value: "square"
    },
    {
        displayName: "Circle",
        value: "circle"
    },
    {
        displayName: "Diamond",
        value: "diamond"
    },
    {
        displayName: "Triangle",
        value: "triangle"
    }
];


class StatusColorsCard
    extends FormattingSettingsCard {


    analyseDurchgefuehrt =
        new formattingSettings.ColorPicker({
            name: "analyseDurchgefuehrt",
            displayName: "Analyse durchgeführt",
            value: {
                value: "#FFA62B"
            }
        });


    analyseFreigegeben =
        new formattingSettings.ColorPicker({
            name: "analyseFreigegeben",
            displayName: "Analyse freigegeben",
            value: {
                value: "#3C9673"
            }
        });


    analyseGereviewed =
        new formattingSettings.ColorPicker({
            name: "analyseGereviewed",
            displayName: "Analyse gereviewed",
            value: {
                value: "#F2D92F"
            }
        });


    analyseRcBatchClosed =
        new formattingSettings.ColorPicker({
            name: "analyseRcBatchClosed",
            displayName: "Analyse r/c + Batch closed",
            value: {
                value: "#7A306C"
            }
        });


    durchzufuehrendeAnalyse =
        new formattingSettings.ColorPicker({
            name: "durchzufuehrendeAnalyse",
            displayName: "Durchzuführende Analyse",
            value: {
                value: "#0000FF"
            }
        });


    investigationOpen =
        new formattingSettings.ColorPicker({
            name: "investigationOpen",
            displayName: "Investigation open",
            value: {
                value: "#C73F1E"
            }
        });


    sampleRejectedCanceled =
        new formattingSettings.ColorPicker({
            name: "sampleRejectedCanceled",
            displayName: "Sample rejected/canceled",
            value: {
                value: "#00181B"
            }
        });


    analyseInBearbeitung =
        new formattingSettings.ColorPicker({
            name: "analyseInBearbeitung",
            displayName: "Analyse in Bearbeitung",
            value: {
                value: "#FFD7A2"
            }
        });


    analyseRejectedCanceled =
        new formattingSettings.ColorPicker({
            name: "analyseRejectedCanceled",
            displayName: "Analyse rejected/canceled",
            value: {
                value: "#FF64E3"
            }
        });


    notMapped =
        new formattingSettings.ColorPicker({
            name: "notMapped",
            displayName: "Not mapped",
            value: {
                value: "#B3B3B3"
            }
        });


    name:
        string =
        "statusColors";


    displayName:
        string =
        "Status colors";


    slices:
        Array<FormattingSettingsSlice> = [

            this.analyseDurchgefuehrt,
            this.analyseFreigegeben,
            this.analyseGereviewed,
            this.analyseRcBatchClosed,
            this.durchzufuehrendeAnalyse,
            this.investigationOpen,
            this.sampleRejectedCanceled,
            this.analyseInBearbeitung,
            this.analyseRejectedCanceled,
            this.notMapped
        ];
}


class StatusShapesCard
    extends FormattingSettingsCard {


    analyseDurchgefuehrtShape =
        new formattingSettings.ItemDropdown({
            name: "analyseDurchgefuehrtShape",
            displayName: "Analyse durchgeführt",
            items: shapeItems,
            value: {
                displayName: "Square",
                value: "square"
            }
        });


    analyseFreigegebenShape =
        new formattingSettings.ItemDropdown({
            name: "analyseFreigegebenShape",
            displayName: "Analyse freigegeben",
            items: shapeItems,
            value: {
                displayName: "Circle",
                value: "circle"
            }
        });


    analyseGereviewedShape =
        new formattingSettings.ItemDropdown({
            name: "analyseGereviewedShape",
            displayName: "Analyse gereviewed",
            items: shapeItems,
            value: {
                displayName: "Diamond",
                value: "diamond"
            }
        });


    analyseRcBatchClosedShape =
        new formattingSettings.ItemDropdown({
            name: "analyseRcBatchClosedShape",
            displayName: "Analyse r/c + Batch closed",
            items: shapeItems,
            value: {
                displayName: "Square",
                value: "square"
            }
        });


    durchzufuehrendeAnalyseShape =
        new formattingSettings.ItemDropdown({
            name: "durchzufuehrendeAnalyseShape",
            displayName: "Durchzuführende Analyse",
            items: shapeItems,
            value: {
                displayName: "Triangle",
                value: "triangle"
            }
        });


    investigationOpenShape =
        new formattingSettings.ItemDropdown({
            name: "investigationOpenShape",
            displayName: "Investigation open",
            items: shapeItems,
            value: {
                displayName: "Triangle",
                value: "triangle"
            }
        });


    sampleRejectedCanceledShape =
        new formattingSettings.ItemDropdown({
            name: "sampleRejectedCanceledShape",
            displayName: "Sample rejected/canceled",
            items: shapeItems,
            value: {
                displayName: "Diamond",
                value: "diamond"
            }
        });


    analyseInBearbeitungShape =
        new formattingSettings.ItemDropdown({
            name: "analyseInBearbeitungShape",
            displayName: "Analyse in Bearbeitung",
            items: shapeItems,
            value: {
                displayName: "Circle",
                value: "circle"
            }
        });


    analyseRejectedCanceledShape =
        new formattingSettings.ItemDropdown({
            name: "analyseRejectedCanceledShape",
            displayName: "Analyse rejected/canceled",
            items: shapeItems,
            value: {
                displayName: "Diamond",
                value: "diamond"
            }
        });


    notMappedShape =
        new formattingSettings.ItemDropdown({
            name: "notMappedShape",
            displayName: "Not mapped",
            items: shapeItems,
            value: {
                displayName: "Square",
                value: "square"
            }
        });


    name:
        string =
        "statusShapes";


    displayName:
        string =
        "Status shapes";


    slices:
        Array<FormattingSettingsSlice> = [

            this.analyseDurchgefuehrtShape,
            this.analyseFreigegebenShape,
            this.analyseGereviewedShape,
            this.analyseRcBatchClosedShape,
            this.durchzufuehrendeAnalyseShape,
            this.investigationOpenShape,
            this.sampleRejectedCanceledShape,
            this.analyseInBearbeitungShape,
            this.analyseRejectedCanceledShape,
            this.notMappedShape
        ];
}


class MatrixAppearanceCard
    extends FormattingSettingsCard {


    markSize =
        new formattingSettings.NumUpDown({
            name: "markSize",
            displayName: "Mark size",
            value: 14
        });


    markGap =
        new formattingSettings.NumUpDown({
            name: "markGap",
            displayName: "Gap between marks",
            value: 4
        });


    rowHeight =
        new formattingSettings.NumUpDown({
            name: "rowHeight",
            displayName: "Row height",
            value: 24
        });


    valueFontFamily =
        new formattingSettings.FontPicker({
            name: "valueFontFamily",
            displayName: "Value font family",
            value: "Segoe UI, wf_segoe-ui_normal, helvetica, arial, sans-serif"
        });


    valueFontSize =
        new formattingSettings.NumUpDown({
            name: "valueFontSize",
            displayName: "Value font size",
            value: 10
        });


    headerFontFamily =
        new formattingSettings.FontPicker({
            name: "headerFontFamily",
            displayName: "Header font family",
            value: "Segoe UI, wf_segoe-ui_normal, helvetica, arial, sans-serif"
        });


    headerFontSize =
        new formattingSettings.NumUpDown({
            name: "headerFontSize",
            displayName: "Header font size",
            value: 10
        });


    valueHeaderFontFamily =
        new formattingSettings.FontPicker({
            name: "valueHeaderFontFamily",
            displayName: "Value header font family",
            value: "Segoe UI, wf_segoe-ui_normal, helvetica, arial, sans-serif"
        });


    valueHeaderFontSize =
        new formattingSettings.NumUpDown({
            name: "valueHeaderFontSize",
            displayName: "Value header font size",
            value: 10
        });


    markShape =
        new formattingSettings.ItemDropdown({
            name: "markShape",
            displayName: "Default mark shape",
            items: shapeItems,
            value: {
                displayName: "Square",
                value: "square"
            }
        });


    autoWidth =
        new formattingSettings.ToggleSwitch({
            name: "autoWidth",
            displayName: "Auto-size columns",
            value: true
        });


    wordWrap =
        new formattingSettings.ToggleSwitch({
            name: "wordWrap",
            displayName: "Word wrap",
            value: true
        });


    showGrid =
        new formattingSettings.ToggleSwitch({
            name: "showGrid",
            displayName: "Show grid lines",
            value: true
        });


    gridColor =
        new formattingSettings.ColorPicker({
            name: "gridColor",
            displayName: "Grid line color",
            value: {
                value: "#E5E5E5"
            }
        });


    headerBackground =
        new formattingSettings.ColorPicker({
            name: "headerBackground",
            displayName: "Header background",
            value: {
                value: "#FFFFFF"
            }
        });


    headerTextColor =
        new formattingSettings.ColorPicker({
            name: "headerTextColor",
            displayName: "Header text color",
            value: {
                value: "#333333"
            }
        });


    rowHeaderBold =
        new formattingSettings.ToggleSwitch({
            name: "rowHeaderBold",
            displayName: "Bold hierarchy labels",
            value: true
        });


    name:
        string =
        "matrixAppearance";


    displayName:
        string =
        "Matrix appearance";


    slices:
        Array<FormattingSettingsSlice> = [

            this.markSize,
            this.markGap,
            this.rowHeight,
            this.valueFontFamily,
            this.valueFontSize,
            this.headerFontFamily,
            this.headerFontSize,
            this.valueHeaderFontFamily,
            this.valueHeaderFontSize,
            this.markShape,
            this.autoWidth,
            this.wordWrap,
            this.showGrid,
            this.gridColor,
            this.headerBackground,
            this.headerTextColor,
            this.rowHeaderBold
        ];
}


class SortingControlsCard
    extends FormattingSettingsCard {


    showColumnSortButtons =
        new formattingSettings.ToggleSwitch({
            name: "showColumnSortButtons",
            displayName: "Show column sort buttons",
            value: true
        });


    showRowSortButtons =
        new formattingSettings.ToggleSwitch({
            name: "showRowSortButtons",
            displayName: "Show row sort buttons",
            value: true
        });


    name:
        string =
        "sortingControls";


    displayName:
        string =
        "Sorting controls";


    slices:
        Array<FormattingSettingsSlice> = [

            this.showColumnSortButtons,
            this.showRowSortButtons
        ];
}


class LegendCard
    extends FormattingSettingsCard {


    showLegend =
        new formattingSettings.ToggleSwitch({
            name: "showLegend",
            displayName: "Show legend",
            value: true
        });


    position =
        new formattingSettings.ItemDropdown({
            name: "position",
            displayName: "Position",
            items: [
                {
                    displayName: "Top",
                    value: "top"
                },
                {
                    displayName: "Bottom",
                    value: "bottom"
                },
                {
                    displayName: "Left",
                    value: "left"
                },
                {
                    displayName: "Right",
                    value: "right"
                }
            ],
            value: {
                displayName: "Top",
                value: "top"
            }
        });


    fontFamily =
        new formattingSettings.FontPicker({
            name: "fontFamily",
            displayName: "Legend font family",
            value: "Segoe UI, wf_segoe-ui_normal, helvetica, arial, sans-serif"
        });


    fontSize =
        new formattingSettings.NumUpDown({
            name: "fontSize",
            displayName: "Font size",
            value: 10
        });


    showTitle =
        new formattingSettings.ToggleSwitch({
            name: "showTitle",
            displayName: "Show title",
            value: true
        });


    titleText =
        new formattingSettings.TextInput({
            name: "titleText",
            displayName: "Title text",
            value: "",
            placeholder: "Auto: category field name"
        });


    titleFontFamily =
        new formattingSettings.FontPicker({
            name: "titleFontFamily",
            displayName: "Title font family",
            value: "Segoe UI, wf_segoe-ui_normal, helvetica, arial, sans-serif"
        });


    titleFontSize =
        new formattingSettings.NumUpDown({
            name: "titleFontSize",
            displayName: "Title font size",
            value: 10
        });


    name:
        string =
        "legend";


    displayName:
        string =
        "Legend";


    slices:
        Array<FormattingSettingsSlice> = [

            this.showLegend,
            this.position,
            this.fontFamily,
            this.fontSize,
            this.showTitle,
            this.titleText,
            this.titleFontFamily,
            this.titleFontSize
        ];
}


export class VisualFormattingSettingsModel
    extends FormattingSettingsModel {


    statusColors =
        new StatusColorsCard();


    statusShapes =
        new StatusShapesCard();


    matrixAppearance =
        new MatrixAppearanceCard();


    sortingControls =
        new SortingControlsCard();


    legend =
        new LegendCard();


    cards = [

        this.statusColors,
        this.statusShapes,
        this.matrixAppearance,
        this.sortingControls,
        this.legend
    ];
}
