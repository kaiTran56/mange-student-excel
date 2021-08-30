package com.tranquyet.util.excel.table.chart;

import com.tranquyet.constant.excel.GraphExcelValueConstant;
import com.tranquyet.service.StudentStorageService;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.Map;

public class ChartTableExcel {

    public static void createChartTable(XSSFSheet sheet, int dynamicPosition) {

        Map<Integer, Long> mapData = StudentStorageService.countSameYear();
        XSSFDrawing drawing = sheet.createDrawingPatriarch();

        // create dynamic value to change position
        XSSFClientAnchor anchor = drawing.createAnchor(2, 2, 0, 0
                , 2
                , dynamicPosition
                , 5
                , 15 + dynamicPosition); // need fix

        XSSFChart chart = drawing.createChart(anchor);
        ChartComponent.setRoundedCorners(chart, false);

        chart.setTitleText("Chart Title");
        chart.setTitleOverlay(false);
        // create legend
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.RIGHT);
        // create bottom axis
        //204-204-204
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle(GraphExcelValueConstant.YEAR_CHART_TABLE);
        ChartComponent.setColorAxis(chart, 0, new byte[]{(byte) 204, (byte) 204, (byte) 204});
        /**
         * left axis
         */
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle(GraphExcelValueConstant.YEAR_CHART_TABLE);
        leftAxis.setCrossBetween(AxisCrossBetween.MIDPOINT_CATEGORY);
        leftAxis.setCrosses(AxisCrosses.MIN); // notice
        ChartComponent.setColorAxis(chart, 1, new byte[]{(byte) 255, (byte) 255, (byte) 255});

        // create right axis
        XDDFValueAxis rightAxis = chart.createValueAxis(AxisPosition.RIGHT);
        rightAxis.setTitle(GraphExcelValueConstant.COUNT_YEAR_CHART_TABLE);
        rightAxis.setCrosses(AxisCrosses.MAX);
        rightAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
        ChartComponent.setColorAxis(chart, 2, new byte[]{(byte) 255, (byte) 255, (byte) 255});
        bottomAxis.crossAxis(rightAxis);
        rightAxis.crossAxis(bottomAxis);

        // get and convert key map to String for category graph
        String[] category = mapData.keySet()
                .stream()
                .map(p -> String.valueOf(p))
                .toArray(String[]::new);

        XDDFDataSource yearsCategory = XDDFDataSourcesFactory.fromArray(category);
        XDDFNumericalDataSource years = XDDFDataSourcesFactory.fromArray(mapData.keySet().toArray(new Integer[0]));
        XDDFNumericalDataSource count = XDDFDataSourcesFactory.fromArray(mapData.values().toArray(new Long[0]));

        XDDFLineChartData dataLine = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, rightAxis);
        XDDFBarChartData data = (XDDFBarChartData) chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
        data.setBarDirection(BarDirection.COL);
        data.setOverlap((byte) 100);

        XDDFBarChartData.Series series1 = (XDDFBarChartData.Series) data.addSeries(yearsCategory, years);
        series1.setTitle(GraphExcelValueConstant.YEAR_CHART_TABLE, null);

        XDDFChartData.Series series3 = dataLine.addSeries(yearsCategory, count);
        series3.setTitle(GraphExcelValueConstant.COUNT_YEAR_CHART_TABLE, null);
        ((XDDFLineChartData.Series) series3).setSmooth(false);
        ((XDDFLineChartData.Series) series3).setMarkerStyle(MarkerStyle.DOT);

        ChartComponent.createLabelGraph(chart);
        chart.plot(dataLine);
        chart.plot(data);

    }
}
