package org.zkoss.zss.ngmodel;

import static org.junit.Assert.assertEquals;

import java.io.*;
import java.net.URL;
import java.util.Locale;

import org.junit.*;
import org.zkoss.util.Locales;
import org.zkoss.zss.ngapi.NImporter;
import org.zkoss.zss.ngapi.impl.imexp.ExcelImportFactory;
import org.zkoss.zss.ngmodel.NChart.NBarDirection;
import org.zkoss.zss.ngmodel.NChart.NChartGrouping;
import org.zkoss.zss.ngmodel.NChart.NChartLegendPosition;
import org.zkoss.zss.ngmodel.NChart.NChartType;
import org.zkoss.zss.ngmodel.chart.NGeneralChartData;

/**
 * @author Hawk
 */
public class ImporterTest extends ImExpTestBase {
	
	private static final URL DEFAULT_CHART_IMPORT_FILE = ImporterTest.class.getResource("book/chart.xlsx");
	private NImporter importer; 
	
	/**
	 * For exporter test to specify its exported file to test.
	 * @param fileUrl
	 */
	static public void setFileUnderTest(URL fileUrl){
		fileForImporterTest = fileUrl;
	}
	
	
	@Before
	public void beforeTest(){
		importer= new ExcelImportFactory().createImporter();
		Locales.setThreadLocal(Locale.TAIWAN);
	}
	
	//API
	
	@Test
	public void importByInputStream(){
		InputStream streamUnderTest = null;
		NBook book = null;
		try {
			book = importer.imports(fileForImporterTest.openStream(), "XSSFBook");
		} catch (IOException e) {
			e.printStackTrace();
		}finally{
			if(streamUnderTest!=null){
				try{
					streamUnderTest.close();
				}catch(Exception e){
					e.printStackTrace();
				}
			}
		}
		
		assertEquals("XSSFBook", book.getBookName());
		assertEquals(7, book.getNumOfSheet());
	}
	
	@Test
	public void importByUrl(){
		NBook book = null;
		try {
			book = importer.imports(fileForImporterTest, "XSSFBook");
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		assertEquals("XSSFBook", book.getBookName());
		assertEquals(7, book.getNumOfSheet());
	}
	
	@Test
	public void importByFile() {
		NBook book = null;
		try {
			book = importer.imports(new File(fileForImporterTest.toURI()), "XSSFBook");
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		assertEquals("XSSFBook", book.getBookName());
		assertEquals(7, book.getNumOfSheet());
	}
	
	//content
	@Test
	public void sheetTest() {
		NBook book = ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook");
		
		sheetTest(book);
	}	
	
	@Test
	public void sheetProtectionTest() {
		NBook book = ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook");
		sheetProtectionTest(book);
	}
	
	@Test
	public void sheetNamedRangeTest() {
		NBook book = ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook");
		sheetNamedRangeTest(book);
	}
	
	@Test
	public void viewInfoTest(){
		NBook book = ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook");
		viewInfoTest(book);
	}
	
	@Test
	public void mergedTest(){
		NBook book = ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook");
		mergedTest(book);
	}
	
	@Test
	public void rowTest(){
		NBook book = ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook");
		rowTest(book);
	}
	
	@Test
	public void columnTest(){
		NBook book = ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook");
		columnTest(book);		
	}
	
	/**
	 * import last column that only has column width change but has all empty cells 
	 */
	@Test
	public void lastChangedColumnTest(){
		NBook book = ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook");
		lastChangedColumnTest(book);
	}

	@Test
	public void cellValueTest() {
		NBook book = ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook");
		
		cellValueTest(book);
	}
	
	@Test
	public void cellStyleTest(){
		NBook book = ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook");
		cellStyleTest(book);
	}

	@Test
	public void cellBorderTest(){
		NBook book = ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook");
		cellBorderTest(book);
	}

	@Test
	public void cellFontNameTest(){
		NBook book = ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook");
		cellFontNameTest(book);
	}
	
	@Test
	public void cellFontStyleTest(){
		NBook book = ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook");
		cellFontStyleTest(book);
	}
	
	@Test
	public void cellFontColorTest(){
		NBook book = ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook");
		cellFontColorTest(book);
		
	}

	/**
	 * Information technology Document description and processing languages 
	 * Office Open XML File Formats Part 1: Fundamentals and Markup LanguageReference  
	 * 18.8.30 numFmt (Number Format) 
	 */
	@Test
	public void cellFormatTest(){
		NBook book = ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook");
		cellFormatTest(book);
	}

	/**
	 * Under different locales, TW and US, should import the same format pattern
	 */
	@Test
	public void formatNotDependLocaleTest(){
		NBook book = ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook");
		NSheet sheet = book.getSheetByName("Format");
		assertEquals("zh_TW", Locales.getCurrent().toString());
		assertEquals("m/d/yyyy", sheet.getCell(1, 4).getCellStyle().getDataFormat());
		
		Locales.setThreadLocal(Locale.US);
		book = ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook");
		sheet = book.getSheetByName("Format");
		assertEquals("en_US", Locales.getCurrent().toString());
		assertEquals("m/d/yyyy", sheet.getCell(1, 4).getCellStyle().getDataFormat());
	}
	
	@Test
	public void barChart(){
		NBook book = ImExpTestUtil.loadBook(DEFAULT_CHART_IMPORT_FILE, "Chart");
		NSheet sheet = book.getSheetByName("Bar");
		NChart barChart = sheet.getChart(0);
		
		assertEquals(NChartType.BAR,barChart.getType());
		assertEquals("Bar Chart Title",barChart.getTitle());
		
		assertEquals(480, barChart.getAnchor().getWidth());
		assertEquals(288, barChart.getAnchor().getHeight());
		
		assertEquals(NBarDirection.HORIZONTAL, barChart.getBarDirection());
		assertEquals(NChartGrouping.CLUSTERED, barChart.getGrouping());
		assertEquals(false, barChart.isThreeD());
		assertEquals(NChartLegendPosition.RIGHT, barChart.getLegendPosition());
		
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
		
		NChart barChart3D = sheet.getChart(1);
		assertEquals(true, barChart3D.isThreeD());
	}
	
	@Test
	public void columnChart(){
		NBook book = ImExpTestUtil.loadBook(DEFAULT_CHART_IMPORT_FILE, "Chart");
		NSheet sheet = book.getSheetByName("Column");
		NChart columnChart = sheet.getChart(0);
		assertEquals(NChartType.COLUMN,columnChart.getType());
		assertEquals(NBarDirection.VERTICAL, columnChart.getBarDirection());
		assertEquals(NChartLegendPosition.TOP, columnChart.getLegendPosition());
		
		NGeneralChartData chartData = (NGeneralChartData)columnChart.getData();
		assertEquals(4, chartData.getNumOfCategory());
	
		NChart column3dChart = sheet.getChart(1);
		assertEquals(NChartGrouping.STACKED, column3dChart.getGrouping());
	}
	
	@Test
	public void areaChart(){
		NBook book = ImExpTestUtil.loadBook(DEFAULT_CHART_IMPORT_FILE, "Chart");
		NSheet sheet = book.getSheetByName("Area");
		NChart areaChart = sheet.getChart(0);
		assertEquals(NChartType.AREA,areaChart.getType());
		
		NGeneralChartData chartData = (NGeneralChartData)areaChart.getData();
		assertEquals(8, chartData.getNumOfCategory());
		
		NChart area3dChart = sheet.getChart(1);
		assertEquals(NChartGrouping.STANDARD, area3dChart.getGrouping());
		assertEquals(NChartLegendPosition.BOTTOM, area3dChart.getLegendPosition());
	}
	
	@Test
	public void bubbleChart(){
		NBook book = ImExpTestUtil.loadBook(DEFAULT_CHART_IMPORT_FILE, "Chart");
		NSheet sheet = book.getSheetByName("Bubble");
		NChart bubbleChart = sheet.getChart(0);
		assertEquals("Sales",bubbleChart.getTitle()); 
		assertEquals(NChartType.BUBBLE, bubbleChart.getType());
		
		NGeneralChartData chartData = (NGeneralChartData)bubbleChart.getData();
		assertEquals(0, chartData.getNumOfCategory());
		assertEquals(1, chartData.getNumOfSeries());
	}
	
	@Test
	public void doughnutChart(){
		NBook book = ImExpTestUtil.loadBook(DEFAULT_CHART_IMPORT_FILE, "Chart");
		NSheet sheet = book.getSheetByName("Doughnut");
		NChart doughnutChart = sheet.getChart(0);
		assertEquals(NChartType.DOUGHNUT, doughnutChart.getType());
		
		NGeneralChartData chartData = (NGeneralChartData)doughnutChart.getData();
		assertEquals(8, chartData.getNumOfCategory());
	}
	
	@Test
	public void lineChart(){
		NBook book = ImExpTestUtil.loadBook(DEFAULT_CHART_IMPORT_FILE, "Chart");
		NSheet sheet = book.getSheetByName("Line");
		NChart lineChart = sheet.getChart(0);
		assertEquals(NChartType.LINE, lineChart.getType());
		NGeneralChartData chartData = (NGeneralChartData)lineChart.getData();
		assertEquals(3, chartData.getNumOfSeries());
		
		NChart line3dChart = sheet.getChart(1);
		assertEquals(true, line3dChart.isThreeD());
		assertEquals("Line 3D Title",line3dChart.getTitle()); 
		chartData = (NGeneralChartData)line3dChart.getData();
		assertEquals(3, chartData.getNumOfSeries());
	}
	
	@Test
	public void pieChart(){
		NBook book = ImExpTestUtil.loadBook(DEFAULT_CHART_IMPORT_FILE, "Chart");
		NSheet sheet = book.getSheetByName("Pie");
		NChart pieChart = sheet.getChart(0);
		assertEquals(NChartType.PIE, pieChart.getType());
		assertEquals(null,pieChart.getTitle());
		NGeneralChartData chartData = (NGeneralChartData)pieChart.getData();
		assertEquals(1, chartData.getNumOfSeries());
		
		NChart pie3dChart = sheet.getChart(1);
		assertEquals(NChartType.PIE, pie3dChart.getType());
		assertEquals(true, pie3dChart.isThreeD());
	}
	
	@Test
	public void scatterChart(){
		NBook book = ImExpTestUtil.loadBook(DEFAULT_CHART_IMPORT_FILE, "Chart");
		NSheet sheet = book.getSheetByName("Scatter");
		NChart scatterChart = sheet.getChart(0);
		assertEquals(NChartType.SCATTER, scatterChart.getType());
		
		NGeneralChartData chartData = (NGeneralChartData)scatterChart.getData();
		assertEquals(3, chartData.getNumOfSeries());
		assertEquals("Internet Explorer", chartData.getSeries(0).getName());
		assertEquals(0.3427, chartData.getSeries(0).getYValue(0));
		assertEquals(0.327, chartData.getSeries(0).getYValue(1));
		assertEquals(0.3168, chartData.getSeries(0).getYValue(2));
	}
	
}



