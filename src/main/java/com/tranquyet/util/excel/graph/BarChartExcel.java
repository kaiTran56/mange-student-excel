package com.tranquyet.util.excel.graph;

import com.tranquyet.constant.excel.GraphExcelValueConstant;
import com.tranquyet.controller.ExcelController;
import com.tranquyet.util.excel.ComponentExcel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;

import java.util.Map;

/**
 * BarChartExcel will create a bar chart base the table which uses to count the number of student having the same year of birth
 */
public class BarChartExcel {

    /**
     * This method would activate <pre>
     *      createCountYearTable: create a table which distinguish the same year of birth of student
     *      createGraph: create a bar chart to describe the table previously
     *  </pre>
     *
     * @param workbook   representation of an Excel workbook
     * @param chartSheet create Chart Sheet
     */
    public static void createGraphSheet(XSSFWorkbook workbook, XSSFSheet chartSheet) {

        ExcelController excelControlle = new ExcelController();
        Map<Integer, Long> mapSameAge = excelControlle.statisticSameYear();

        // draw the table
        createCountYearTable(workbook, chartSheet, mapSameAge);
        // draw graph
        createGraphCheck(chartSheet, mapSameAge.size());

        createBreakPage(chartSheet);

    }

    /**
     * create a bar chart to describe the count year table previously
     *
     * @param chartSheet     create Chart Sheet
     * @param dynamicPos a dynamic variable to expand the width of chart and border
     *                       when sizeMap increase then the width of chart and border will increase a quantity respectively
     */
    public static void createGraphCheck(XSSFSheet chartSheet, int dynamicPos) {

        XSSFDrawing drawing = chartSheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(2, 2, 0, 0
                , GraphExcelValueConstant.START_WIDTH_ANCHOR_POSITION
                , GraphExcelValueConstant.START_HEIGHT_ANCHOR_POSITION
                , GraphExcelValueConstant.END_WIDTH_ANCHOR_POSITION
                , GraphExcelValueConstant.END_HEIGHT_ANCHOR_POSITION);

        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText("Chart Title");


        setRoundedCorners(chart, false);
        CTChart ctChart = chart.getCTChart();
        CTPlotArea ctPlotArea = ctChart.getPlotArea();

        //the bar chart
        CTBarChart ctBarChart = ctPlotArea.addNewBarChart();
        CTBoolean ctBoolean = ctBarChart.addNewVaryColors();
        ctBoolean.setVal(true);
        ctBarChart.addNewBarDir().setVal(STBarDir.COL);

//the bar series
        StringBuilder positionYearCell = new StringBuilder().append(GraphExcelValueConstant.CHART_SHEET_EXCEL).append("!$K$5:$K").append("$"+(4+dynamicPos));
        CTBarSer ctBarSer = ctBarChart.addNewSer();
        CTSerTx ctSerTx = ctBarSer.addNewTx();
        CTStrRef ctStrRef = ctSerTx.addNewStrRef();
        ctStrRef.setF(GraphExcelValueConstant.CHART_SHEET_EXCEL+"!$K$4");
        ctBarSer.addNewIdx().setVal(0);
        CTAxDataSource cttAxDataSource = ctBarSer.addNewCat();
        ctStrRef = cttAxDataSource.addNewStrRef();
        // ctStrRef.setF("!$K$5:$K$9"); // set name bottom // dynamic
        ctStrRef.setF(positionYearCell.toString());

        CTNumDataSource ctNumDataSource = ctBarSer.addNewVal();
        CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
        ctNumRef.setF(positionYearCell.toString()); // set value for bar char // dynamic

        ctBarSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[]{0, 0, 0});

        //telling the BarChart that it has axes and giving them Ids
        ctBarChart.addNewAxId().setVal(123456); //cat axis 1 (bars)
        ctBarChart.addNewAxId().setVal(123457); //val axis 1 (left)

        //the line chart
        CTLineChart ctLineChart = ctPlotArea.addNewLineChart();
        ctBoolean = ctLineChart.addNewVaryColors();
        ctBoolean.setVal(true);
        ctLineChart.addNewSmooth().setVal(true);
        ctLineChart.addNewSmooth().setVal(false);


        //the line series
        StringBuilder positionCountCell = new StringBuilder().append(GraphExcelValueConstant.CHART_SHEET_EXCEL)
                .append("!$L$5:$L").append("$"+(4+dynamicPos));
        CTLineSer ctLineSer = ctLineChart.addNewSer();
        ctLineSer.addNewSmooth().setVal(false);
        ctSerTx = ctLineSer.addNewTx();
        ctStrRef = ctSerTx.addNewStrRef();
        ctStrRef.setF(GraphExcelValueConstant.CHART_SHEET_EXCEL+"!$L$4");// set value
        ctLineSer.addNewIdx().setVal(1);
        cttAxDataSource = ctLineSer.addNewCat();
        ctStrRef = cttAxDataSource.addNewStrRef();
        ctStrRef.setF(positionYearCell.toString()); // set name column
        ctNumDataSource = ctLineSer.addNewVal();
        ctNumRef = ctNumDataSource.addNewNumRef();
        ctNumRef.setF(positionCountCell.toString()); // set value

        //at least the border lines in Libreoffice Calc ;-)
        ctLineSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[]{(byte) 255, (byte) 153, 0});

        //telling the LineChart that it has axes and giving them Ids
        ctLineChart.addNewAxId().setVal(123458); //cat axis 2 (lines)
        ctLineChart.addNewAxId().setVal(123459); //val axis 2 (right)
        //cat axis 1 (bars)
        CTCatAx ctCatAx = ctPlotArea.addNewCatAx();
        ctCatAx.addNewAxId().setVal(123456); //id of the cat axis
        ctCatAx.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[]{(byte) 204, (byte) 204, (byte) 204});
        CTScaling ctScaling = ctCatAx.addNewScaling();

        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctCatAx.addNewDelete().setVal(false);
        ctCatAx.addNewAxPos().setVal(STAxPos.B);
        ctCatAx.addNewCrossAx().setVal(123457); //id of the val axis
        ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

        //val axis 1 (left)
        CTValAx ctValAx = ctPlotArea.addNewValAx();
        ctValAx.addNewAxId().setVal(123457); //id of the val axis
        ;
        ctScaling = ctValAx.addNewScaling();
        ctValAx.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[]{(byte) 255, (byte) 255, (byte) 255});
        ctValAx.addNewDelete().setVal(false);
        ctValAx.addNewAxPos().setVal(STAxPos.L);
        ctValAx.addNewCrossAx().setVal(123456); //id of the cat axis
        ctValAx.addNewCrosses().setVal(STCrosses.AUTO_ZERO); //this val axis crosses the cat axis at zero
        ctValAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

        //cat axis 2 (lines)
        ctCatAx = ctPlotArea.addNewCatAx();
        ctCatAx.addNewAxId().setVal(123458); //id of the cat axis
        ctScaling = ctCatAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctCatAx.addNewDelete().setVal(true); //this cat axis is deleted
        ctCatAx.addNewAxPos().setVal(STAxPos.B);
        ctCatAx.addNewCrossAx().setVal(123459); //id of the val axis
        ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

        //val axis 2 (right)
        ctValAx = ctPlotArea.addNewValAx();
        ctValAx.addNewAxId().setVal(123459); //id of the val axis
        ctValAx.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[]{(byte) 255, (byte) 255, (byte) 255});
        ctScaling = ctValAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctValAx.addNewDelete().setVal(false);
        ctValAx.addNewAxPos().setVal(STAxPos.R);
        ctValAx.addNewCrossAx().setVal(123458); //id of the cat axis
        ctValAx.addNewCrosses().setVal(STCrosses.MAX); //this val axis crosses the cat axis at max value
        ctValAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

        /**
         * set title for axis
         */
        chart.getAxes().get(0).setTitle(GraphExcelValueConstant.YEAR_CHART_TABLE); // bottom axis
        chart.getAxes().get(3).setTitle(GraphExcelValueConstant.COUNT_YEAR_CHART_TABLE); // right axis
        chart.getAxes().get(2).setTitle(GraphExcelValueConstant.YEAR_CHART_TABLE); // left axis

        /**
         * format label for barchart
         */
        chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).addNewDLbls();
        chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowVal().setVal(true);
        chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowLegendKey().setVal(false);
        chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowCatName().setVal(false);
        chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowSerName().setVal(false);

        /**
         * format label for line graph
         */
        chart.getCTChart().getPlotArea().getLineChartArray(0).getSerArray(0).addNewDLbls();
        chart.getCTChart().getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls()
                .addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(new byte[]{(byte) 255, (byte) 255, (byte) 255});
        chart.getCTChart().getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowVal().setVal(true);
        chart.getCTChart().getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowLegendKey().setVal(false);
        chart.getCTChart().getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowCatName().setVal(false);
        chart.getCTChart().getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowSerName().setVal(false);

        /**
         * Set position Legend
         */
        CTLegend ctLegend = ctChart.addNewLegend();
        ctLegend.addNewLegendPos().setVal(STLegendPos.R);
        ctLegend.addNewOverlay().setVal(false);
    }

    /**
     * create a table which distinguish the same year of birth of student
     *
     * @param workbook   representation of an Excel workbook
     * @param chartSheet create Chart Sheet
     * @param mapSameAge a set of data which comprise key - unique year and value - the number of student having the same year
     */
    public static void createCountYearTable(Workbook workbook, Sheet chartSheet, Map<Integer, Long> mapSameAge) {
        ComponentExcel componentExcel = ComponentExcel.getInstance();
        Row row = null;
        Cell cell = null;

        /**
         * create Year field
         */
        row = chartSheet.createRow(3);
        cell = row.createCell(10);
        cell.setCellValue(GraphExcelValueConstant.YEAR_CHART_TABLE);
        componentExcel.createDateType(cell, workbook);

        /**
         * create count field
         */
        cell = row.createCell(11);
        cell.setCellValue(GraphExcelValueConstant.COUNT_YEAR_CHART_TABLE);
        componentExcel.createDateType(cell, workbook);

        chartSheet.setColumnWidth(10, 14 * 256);
        chartSheet.setColumnWidth(11, 15 * 256);
        // create cell and set value
        int increase = 1;
        for (Integer year : mapSameAge.keySet()) {
            row = chartSheet.createRow(3 + increase);
            cell = row.createCell(10);
            cell.setCellValue(year);
            componentExcel.createDateType(cell, workbook);
            cell = row.createCell(11);
            cell.setCellValue(mapSameAge.get(year));
            componentExcel.createDateType(cell, workbook);
            increase++;
        }


    }

    private static void setRoundedCorners(XSSFChart chart, boolean setVal) {
        if (chart.getCTChartSpace().getRoundedCorners() == null) chart.getCTChartSpace().addNewRoundedCorners();
        chart.getCTChartSpace().getRoundedCorners().setVal(setVal);
    }

    public static void createBreakPage(Sheet chartSheet) {
        chartSheet.setRowBreak(16);
        chartSheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
        chartSheet.setFitToPage(true);
        PrintSetup printSetup = chartSheet.getPrintSetup();
        printSetup.setFitHeight((short) 0);
        printSetup.setFitWidth((short) 1);
    }

}
