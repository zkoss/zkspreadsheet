package org.zkoss.zss.model;

import static org.junit.Assert.assertEquals;

import org.junit.*;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SChart;
import org.zkoss.zss.model.SPicture;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SChart.*;
import org.zkoss.zss.model.SPicture.Format;
import org.zkoss.zss.model.chart.SGeneralChartData;

public class ImporterXlsTest extends ImporterTest {
	@BeforeClass
	static public void beforeClass() {
		Setup.touch();
	}
	@Before
	public void setupTestFile(){
		IMPORT_FILE_UNDER_TEST = ImporterTest.class.getResource("book/import.xls");
		CHART_IMPORT_FILE_UNDER_TEST = ImporterTest.class.getResource("book/chart.xls");
		PICTURE_IMPORT_FILE_UNDER_TEST = ImporterTest.class.getResource("book/picture.xls");
	}

	@Override
	public void sheetTest() {
		super.sheetTest();
	}
	
	@Override
	public void areaChart() {
		SBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		SSheet sheet = book.getSheetByName("Area");
		SChart areaChart = sheet.getChart(0);
		assertEquals(ChartType.AREA,areaChart.getType());
		
		//a chart locating in one column and one row test
		assertEquals(493, areaChart.getAnchor().getWidth());
		assertEquals(283, areaChart.getAnchor().getHeight());
		
		SGeneralChartData chartData = (SGeneralChartData)areaChart.getData();
		assertEquals(8, chartData.getNumOfCategory());
		
		SChart area3dChart = sheet.getChart(1);
		assertEquals(571, area3dChart.getAnchor().getWidth());
	}
	
	@Override
	public void barChart() {
		SBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		SSheet sheet = book.getSheetByName("Bar");
		SChart barChart = sheet.getChart(0);
		
		assertEquals(ChartType.BAR,barChart.getType());
		
		assertEquals(480, barChart.getAnchor().getWidth());
		assertEquals(284, barChart.getAnchor().getHeight());
		assertEquals(25, barChart.getAnchor().getXOffset());
		assertEquals(7, barChart.getAnchor().getYOffset());
		
		assertEquals(false, barChart.isThreeD());
		
		//data
		SGeneralChartData chartData = (SGeneralChartData)barChart.getData();
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
		
		SChart barChart3D = sheet.getChart(1);
		assertEquals(true, barChart3D.isThreeD());
	}
	
	@Override
	public void bubbleChart() {
		SBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		bubbleChart(book);
	}
	
	@Override
	public void columnChart() {
		SBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		SSheet sheet = book.getSheetByName("Column");
		SChart columnChart = sheet.getChart(0);
		assertEquals(ChartType.COLUMN,columnChart.getType());
		
		SGeneralChartData chartData = (SGeneralChartData)columnChart.getData();
		assertEquals(4, chartData.getNumOfCategory());
		
		SChart column3dChart = sheet.getChart(1);
		SGeneralChartData chart3dData = (SGeneralChartData)column3dChart.getData();
		assertEquals(4, chart3dData.getNumOfCategory());
	}
	
	@Override
	public void doughnutChart() {
		SBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		doughnutChart(book);
	}
	
	@Override
	public void lineChart() {
		SBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		lineChart(book);
	}
	
	@Override
	public void pieChart() {
		SBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		pieChart(book);
	}
	
	@Override
	public void scatterChart() {
		SBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		scatterChart(book);
	}
	
	@Override
	public void picture(){
		SBook book = ImExpTestUtil.loadBook(PICTURE_IMPORT_FILE_UNDER_TEST, "Picture");
		picture(book);
		
		SSheet sheet2 = book.getSheet(1);
		assertEquals(2,sheet2.getPictures().size());
		
		SPicture flowerJpg = sheet2.getPicture(0);
		assertEquals(Format.JPG, flowerJpg.getFormat());
		assertEquals(569, flowerJpg.getAnchor().getWidth());
		assertEquals(427, flowerJpg.getAnchor().getHeight());
		
		//different spec in XLS, GIF pictures are treated as PNG
		SPicture rainbowGif = sheet2.getPicture(1);
		assertEquals(Format.PNG, rainbowGif.getFormat());
		assertEquals(613, rainbowGif.getAnchor().getWidth());
		assertEquals(345, rainbowGif.getAnchor().getHeight());
	}
	
	@Override @Ignore("not support")
	public void validation() {
	}
	
	@Override @Ignore("not support")
	public void autoFilter() {
	}
}
