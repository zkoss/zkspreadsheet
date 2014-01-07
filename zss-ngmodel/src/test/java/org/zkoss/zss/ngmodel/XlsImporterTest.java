package org.zkoss.zss.ngmodel;

import static org.junit.Assert.assertEquals;

import org.junit.*;
import org.zkoss.zss.ngmodel.NChart.*;
import org.zkoss.zss.ngmodel.chart.NGeneralChartData;

public class XlsImporterTest extends ImporterTest {

	@Before
	public void setupTestFile(){
		IMPORT_FILE_UNDER_TEST = ImporterTest.class.getResource("book/import.xls");
		CHART_IMPORT_FILE_UNDER_TEST = ImporterTest.class.getResource("book/chart.xls");
	}
	
	@Override
	public void areaChart() {
		NBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		NSheet sheet = book.getSheetByName("Area");
		NChart areaChart = sheet.getChart(0);
		assertEquals(NChartType.AREA,areaChart.getType());
		
		//a chart locating in one column and one row test
		assertEquals(493, areaChart.getAnchor().getWidth());
		assertEquals(283, areaChart.getAnchor().getHeight());
		
		NGeneralChartData chartData = (NGeneralChartData)areaChart.getData();
		assertEquals(8, chartData.getNumOfCategory());
		
		NChart area3dChart = sheet.getChart(1);
		assertEquals(571, area3dChart.getAnchor().getWidth());
	}
	
	@Override
	public void barChart() {
		NBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		NSheet sheet = book.getSheetByName("Bar");
		NChart barChart = sheet.getChart(0);
		
		assertEquals(NChartType.BAR,barChart.getType());
		
		assertEquals(480, barChart.getAnchor().getWidth());
		assertEquals(284, barChart.getAnchor().getHeight());
		assertEquals(25, barChart.getAnchor().getXOffset());
		assertEquals(7, barChart.getAnchor().getYOffset());
		
		assertEquals(false, barChart.isThreeD());
		
		//data
		NGeneralChartData chartData = (NGeneralChartData)barChart.getData();
		assertEquals(3, chartData.getNumOfCategory());
		assertEquals("Internet Explorer", chartData.getCategory(0));
		assertEquals("Chrome", chartData.getCategory(1));
		assertEquals("Firefox", chartData.getCategory(2));
		assertEquals(3, chartData.getNumOfSeries());
		assertEquals("January 2012", chartData.getSeries(0).getName());
		assertEquals(0.3427, chartData.getSeries(0).getValue(0));
		assertEquals(0.2599, chartData.getSeries(0).getValue(1));
		assertEquals(0.2268, chartData.getSeries(0).getValue(2));
		assertEquals("February 2012", chartData.getSeries(1).getName());
		assertEquals(0.327, chartData.getSeries(1).getValue(0));
		assertEquals(0.2724, chartData.getSeries(1).getValue(1));
		assertEquals(0.2276, chartData.getSeries(1).getValue(2));
		assertEquals("March 2012", chartData.getSeries(2).getName());
		assertEquals(0.3168, chartData.getSeries(2).getValue(0));
		assertEquals(0.2809, chartData.getSeries(2).getValue(1));
		assertEquals(0.2273, chartData.getSeries(2).getValue(2));
		
		NChart barChart3D = sheet.getChart(1);
		assertEquals(true, barChart3D.isThreeD());
	}
	
	@Override
	public void bubbleChart() {
		NBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		NSheet sheet = book.getSheetByName("Bubble");
		NChart bubbleChart = sheet.getChart(0);
		//bubble chart is Scatter chart in XLS
		assertEquals(NChartType.SCATTER, bubbleChart.getType());
		
		NGeneralChartData chartData = (NGeneralChartData)bubbleChart.getData();
		assertEquals(0, chartData.getNumOfCategory());
		assertEquals(1, chartData.getNumOfSeries());
		assertEquals("String Literal Title", chartData.getSeries(0).getName());
	}
	
	@Override
	public void columnChart() {
		NBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		NSheet sheet = book.getSheetByName("Column");
		NChart columnChart = sheet.getChart(0);
		assertEquals(NChartType.BAR,columnChart.getType());
		
		NGeneralChartData chartData = (NGeneralChartData)columnChart.getData();
		assertEquals(4, chartData.getNumOfCategory());
		
		NChart column3dChart = sheet.getChart(1);
		NGeneralChartData chart3dData = (NGeneralChartData)column3dChart.getData();
		assertEquals(4, chart3dData.getNumOfCategory());
	}
	
	@Override
	public void doughnutChart() {
		NBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		NSheet sheet = book.getSheetByName("Doughnut");
		NChart doughnutChart = sheet.getChart(0);
		assertEquals(NChartType.PIE, doughnutChart.getType());
		
		NGeneralChartData chartData = (NGeneralChartData)doughnutChart.getData();
		assertEquals(8, chartData.getNumOfCategory());
	}
	
	@Override
	public void lineChart() {
		NBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		lineChart(book);
	}
	
	@Override
	public void pieChart() {
		NBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		pieChart(book);
	}
	
	@Override
	public void scatterChart() {
		NBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		scatterChart(book);
	}
}
