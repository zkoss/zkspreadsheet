package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;

import java.io.*;

import org.zkoss.image.AImage;
import org.zkoss.zss.api.*;
import org.zkoss.zss.api.model.*;
import org.zkoss.zss.api.model.Chart.Grouping;
import org.zkoss.zss.api.model.Chart.LegendPosition;
import org.zkoss.zss.api.model.impl.ChartImpl;
import org.zkoss.zss.ngmodel.*;
import org.zkoss.zss.ngmodel.NChart.NChartType;

/**
 * all method implementation to test chart & picture operation
 * @author kuro
 */
public class ChartPictureTestBase {
	
	protected void testDeletePicture(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Picture picture = SheetOperationUtil.addPicture(Ranges.range(sheet), new AImage(new File(ChartPictureTestBase.class.getResource("").getPath() + "book/zklogo.png")));
		assertEquals(1, sheet.getPictures().size());
		SheetOperationUtil.deletePicture(Ranges.range(sheet), picture);
		assertEquals(0, sheet.getPictures().size());
	}
	
	protected void testMovePicture(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Picture picture = SheetOperationUtil.addPicture(Ranges.range(sheet), new AImage(new File(ChartPictureTestBase.class.getResource("").getPath() + "book/zklogo.png")));
		assertEquals(1, sheet.getPictures().size());
		SheetOperationUtil.movePicture(Ranges.range(sheet), picture, 10, 30);
		assertEquals(10 , picture.getAnchor().getRow());
		assertEquals(30 , picture.getAnchor().getColumn());
	}
	
	protected void testAddPicture(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Picture picture = SheetOperationUtil.addPicture(Ranges.range(sheet), new AImage(new File(ChartPictureTestBase.class.getResource("").getPath() + "book/zklogo.png")));
		assertEquals(1, sheet.getPictures().size());
		assertEquals(0 , picture.getAnchor().getRow());
		assertEquals(0 , picture.getAnchor().getColumn());
	}
	
	protected void testMoveChart(Book workbook){
		Sheet sheet = workbook.getSheet("chart-image");
		Chart chart = SheetOperationUtil.addChart(Ranges.range(sheet, 4,1,14,1), Chart.Type.LINE, Grouping.STANDARD, LegendPosition.TOP);
		SheetOperationUtil.moveChart(Ranges.range(sheet), chart, 10, 20);
		assertEquals(10 , chart.getAnchor().getRow());
		assertEquals(20 , chart.getAnchor().getColumn());
	}
	
	protected void testDeleteChart(Book workbook){
		Sheet sheet = workbook.getSheet("chart-image");
		Chart chart = SheetOperationUtil.addChart(Ranges.range(sheet, 4,1,14,1), Chart.Type.LINE, Grouping.STANDARD, LegendPosition.TOP);
		assertEquals(1, sheet.getCharts().size());
		
		SheetOperationUtil.deleteChart(Ranges.range(sheet), chart);
		assertEquals(0, sheet.getCharts().size());
	}
	
	protected void testAddLineChart(Book workbook){
		Sheet sheet = workbook.getSheet("chart-image");
		Chart chart = SheetOperationUtil.addChart(Ranges.range(sheet, 4,1,14,1),  Chart.Type.LINE, Grouping.STANDARD, LegendPosition.TOP);
		assertEquals(1, sheet.getCharts().size());
		NChart nchart = ((ChartImpl)chart).getNative();
		assertEquals(nchart.getType(), NChartType.LINE);
	}
	
	protected void testAddBarChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		SheetOperationUtil.addChart(Ranges.range(sheet, 4,3,14,3), Chart.Type.BAR, Grouping.STANDARD, LegendPosition.TOP);
		assertEquals(1, sheet.getCharts().size());
	}
	
	protected void testAddAreaChart(Book workbook) {
		Sheet sheet = workbook.getSheet("chart-image");
		SheetOperationUtil.addChart(Ranges.range(sheet, 4,3,14,3), Chart.Type.AREA, Grouping.STANDARD, LegendPosition.TOP);
		assertEquals(1, sheet.getCharts().size());
	}
	
	// unsupported 
	protected void testAddBubbleChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		SheetOperationUtil.addChart(Ranges.range(sheet, 4,3,14,3), Chart.Type.BUBBLE, Grouping.STANDARD, LegendPosition.TOP);
		assertEquals(1, sheet.getCharts().size());
	}
	
	protected void testAddColumnChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		SheetOperationUtil.addChart(Ranges.range(sheet, 4,3,14,3), Chart.Type.COLUMN, Grouping.STANDARD, LegendPosition.TOP);
		assertEquals(1, sheet.getCharts().size());
	}
	
	protected void testAddDoughnutChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		SheetOperationUtil.addChart(Ranges.range(sheet, 4,3,14,3), Chart.Type.DOUGHNUT, Grouping.STANDARD, LegendPosition.TOP);
	}
	
	protected void testAddPieChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		SheetOperationUtil.addChart(Ranges.range(sheet, 4,3,14,3), Chart.Type.PIE, Grouping.STANDARD, LegendPosition.TOP);
		assertEquals(1, sheet.getCharts().size());
	}
	
	// Not Implement
	protected void testAddRadarChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		SheetOperationUtil.addChart(Ranges.range(sheet, 4,3,14,3), Chart.Type.RADAR, Grouping.STANDARD, LegendPosition.TOP);
	}
	
	protected void testAddScatterChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		SheetOperationUtil.addChart(Ranges.range(sheet, 4,3,14,3),  Chart.Type.SCATTER, Grouping.STANDARD, LegendPosition.TOP);
		assertEquals(1, sheet.getCharts().size());
	}
	
	// unsupported
	protected void testAddStockChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		SheetOperationUtil.addChart(Ranges.range(sheet, 4,3,14,3), Chart.Type.STOCK, Grouping.STANDARD, LegendPosition.TOP);
	}
	
	// unsupported 
	protected void testAddSurfaceChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		SheetOperationUtil.addChart(Ranges.range(sheet, 4,3,14,3), Chart.Type.SURFACE, Grouping.STANDARD, LegendPosition.TOP);
	}
}
