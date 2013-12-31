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
	public void areaChart(){
		NBook book = ImExpTestUtil.loadBook(DEFAULT_CHART_IMPORT_FILE, "Chart");
		areaChart(book);
	}
	
	@Test
	public void barChart(){
		NBook book = ImExpTestUtil.loadBook(DEFAULT_CHART_IMPORT_FILE, "Chart");
		barChart(book);
		NSheet sheet = book.getSheetByName("Bar");
		NChart barChart = sheet.getChart(0);
		
		assertEquals("Bar Chart Title",barChart.getTitle());
		assertEquals(480, barChart.getAnchor().getWidth());
		assertEquals(288, barChart.getAnchor().getHeight());
	}
	
	@Test
	public void bubbleChtart(){
		NBook book = ImExpTestUtil.loadBook(DEFAULT_CHART_IMPORT_FILE, "Chart");
		bubbleChart(book);
		NSheet sheet = book.getSheetByName("Bubble");
		NChart bubbleChart = sheet.getChart(0);
		assertEquals("Sales",bubbleChart.getTitle()); 
	}
	
	@Test
	public void columnChart(){
		NBook book = ImExpTestUtil.loadBook(DEFAULT_CHART_IMPORT_FILE, "Chart");
		columnChart(book);
	}
	
	@Test
	public void doughnutChart(){
		NBook book = ImExpTestUtil.loadBook(DEFAULT_CHART_IMPORT_FILE, "Chart");
		doughnutChart(book);
	}
	
	@Test
	public void lineChart(){
		NBook book = ImExpTestUtil.loadBook(DEFAULT_CHART_IMPORT_FILE, "Chart");
		lineChart(book);
		NSheet sheet = book.getSheetByName("Line");
		NChart line3dChart = sheet.getChart(1);
		assertEquals("Line 3D Title",line3dChart.getTitle()); 
	}
	
	@Test
	public void pieChart(){
		NBook book = ImExpTestUtil.loadBook(DEFAULT_CHART_IMPORT_FILE, "Chart");
		pieChart(book);
	}
	
	@Test
	public void scatterChart(){
		NBook book = ImExpTestUtil.loadBook(DEFAULT_CHART_IMPORT_FILE, "Chart");
		scatterChart(book);
	}
	
	@Ignore("XSSFStockChartData implementation is incorrect")
	public void stockChart(){
		NBook book = ImExpTestUtil.loadBook(DEFAULT_CHART_IMPORT_FILE, "Chart");
		NSheet sheet = book.getSheetByName("Stock");
		NChart stockChart = sheet.getChart(0);
		assertEquals(NChartType.STOCK, stockChart.getType());
		
		NGeneralChartData chartData = (NGeneralChartData)stockChart.getData();
		assertEquals(4, chartData.getNumOfSeries());
		assertEquals("Open", chartData.getSeries(0).getName());
		assertEquals("High", chartData.getSeries(1).getName());
		assertEquals("Low", chartData.getSeries(2).getName());
		assertEquals("Close", chartData.getSeries(3).getName());
	}
}



