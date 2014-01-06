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
		
//		assertEquals(NChartGrouping.STANDARD, area3dChart.getGrouping());
//		assertEquals(NChartLegendPosition.BOTTOM, area3dChart.getLegendPosition());
	}
	
	@Override
	public void barChart() {
		//TODO
	}
	
	@Override
	public void bubbleChtart() {
		//TODO
	}
	
	@Override
	public void columnChart() {
		// TODO Auto-generated method stub
	}
	
	@Override
	public void doughnutChart() {
		// TODO Auto-generated method stub
	}
	
	@Override
	public void lineChart() {
		// TODO Auto-generated method stub
	}
	
	@Override
	public void pieChart() {
		// TODO Auto-generated method stub
	}
	
	@Override
	public void scatterChart() {
		// TODO Auto-generated method stub
	}
}
