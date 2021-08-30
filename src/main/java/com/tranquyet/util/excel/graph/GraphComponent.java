package com.tranquyet.util.excel.graph;

import com.tranquyet.constant.excel.GraphExcelValueConstant;
import com.tranquyet.constant.excel.TableExcelValueConstant;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLegend;
import org.openxmlformats.schemas.drawingml.x2006.chart.STLegendPos;

/**
 * GraphComponent uses to create style for cell, data and graph in  the Chart Sheet
 */
public class GraphComponent {

    /**
     * automatically create page break for sheet
     * @param chartSheet representation of an Excel sheet
     */
    protected static void createBreakPage(Sheet chartSheet) {
        chartSheet.setRowBreak(16);
        chartSheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
        chartSheet.setFitToPage(true);
        PrintSetup printSetup = chartSheet.getPrintSetup();
        printSetup.setFitHeight((short) 0);
        printSetup.setFitWidth((short) 1);
    }

    /**
     * make the border of anchor square
     * @param chart current chart
     * @param setVal check status of border square or not
     */
    protected static void setRoundedCorners(XSSFChart chart, boolean setVal) {
        if (chart.getCTChartSpace().getRoundedCorners() == null) chart.getCTChartSpace().addNewRoundedCorners();
        chart.getCTChartSpace().getRoundedCorners().setVal(setVal);
    }

    /**
     * create label for chart and set position for legend of chart
     * @param chart current chart need
     * @param ctChart current chart but a child of chart and uses to set label for chart
     */
    protected static void createLegendAndLabel(XSSFChart chart, CTChart ctChart) {
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
     *create style for data
     * @param cell current cell of Excel File
     * @param wb representation of an Excel workbook
     */
    protected static void createDataType(Cell cell, Workbook wb) {

        Font font = wb.createFont();
        font.setFontHeightInPoints(TableExcelValueConstant.SIZE_WORD_EXCEL);
        font.setFontName(TableExcelValueConstant.FONT_STYLE_EXCEL);
        font.setBold(TableExcelValueConstant.BOLD_WORD_EXCEL);

        CellStyle style = wb.createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellStyle(style);
    }
}
