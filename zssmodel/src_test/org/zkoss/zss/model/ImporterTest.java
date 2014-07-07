package org.zkoss.zss.model;

import static org.junit.Assert.*;

import java.io.*;
import java.net.URL;
import java.util.Locale;

import org.junit.*;
import org.zkoss.util.Locales;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SChart;
import org.zkoss.zss.model.SPicture;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SChart.ChartType;
import org.zkoss.zss.model.SPicture.Format;
import org.zkoss.zss.model.chart.SGeneralChartData;
import org.zkoss.zss.range.SImporter;
import org.zkoss.zss.range.impl.imexp.ExcelImportFactory;

/**
 * @author Hawk
 */
public class ImporterTest extends ImExpTestBase {
	
	private SImporter importer;
	 
	@BeforeClass
	static public void beforeClass() {
		Setup.touch();
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
		SBook book = null;
		try {
			book = importer.imports(IMPORT_FILE_UNDER_TEST.openStream(), "XSSFBook");
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
		assertEquals(10, book.getNumOfSheet());
	}
	
	@Test
	public void importByUrl(){
		SBook book = null;
		try {
			book = importer.imports(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		assertEquals("XSSFBook", book.getBookName());
		assertEquals(10, book.getNumOfSheet());
	}
	
	@Test
	public void importByFile() {
		SBook book = null;
		try {
			book = importer.imports(new File(IMPORT_FILE_UNDER_TEST.toURI()), "XSSFBook");
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		assertEquals("XSSFBook", book.getBookName());
		assertEquals(10, book.getNumOfSheet());
	}
	
	//content
	@Test
	public void sheetTest() {
		SBook book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		
		sheetTest(book);
	}	
	
	@Test
	public void sheetProtectionTest() {
		SBook book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		sheetProtectionTest(book);
	}
	
	@Test
	public void sheetNamedRangeTest() {
		SBook book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		sheetNamedRangeTest(book);
	}
	
	@Test
	public void viewInfoTest(){
		SBook book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		viewInfoTest(book);
	}
	
	@Test
	public void mergedTest(){
		SBook book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		mergedTest(book);
	}
	
	@Test
	public void rowTest(){
		SBook book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		rowTest(book);
	}
	
	@Test
	public void columnTest(){
		SBook book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		columnTest(book);		
	}
	
	/**
	 * import last column that only has column width change but has all empty cells 
	 */
	@Test
	public void lastChangedColumnTest(){
		SBook book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		lastChangedColumnTest(book);
	}

	@Test
	public void cellValueTest() {
		SBook book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		
		cellValueTest(book);
	}
	
	@Test
	public void cellStyleTest(){
		SBook book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		cellStyleTest(book);
	}

	@Test
	public void cellBorderTest(){
		SBook book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		cellBorderTest(book);
	}

	@Test
	public void cellFontNameTest(){
		SBook book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		cellFontNameTest(book);
	}
	
	@Test
	public void cellFontStyleTest(){
		SBook book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		cellFontStyleTest(book);
	}
	
	@Test
	public void cellFontColorTest(){
		SBook book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		cellFontColorTest(book);
		
	}

	/**
	 * Information technology Document description and processing languages 
	 * Office Open XML File Formats Part 1: Fundamentals and Markup LanguageReference  
	 * 18.8.30 numFmt (Number Format) 
	 */
	@Test
	public void cellFormatTest(){
		SBook book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		cellFormatTest(book);
	}
	
	/**
	 * Under different locales, TW and US and DE, should import the same formula
	 */
	@Test
	public void formulaNotDependLocaleTest(){
		//Locale.TAIWAN
		SBook book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		SSheet sheet = book.getSheetByName("Value");
		assertEquals("zh_TW", Locales.getCurrent().toString());
		
		//formula
		assertEquals(SCell.CellType.FORMULA, sheet.getCell(3,1).getType());
		assertEquals("SUM(10,20)", sheet.getCell(3,1).getFormulaValue());
		assertEquals("ISBLANK(B1)", sheet.getCell(3,2).getFormulaValue());
		assertEquals("B1", sheet.getCell(3,3).getFormulaValue());
		
		//Locale.US
		Locales.setThreadLocal(Locale.US);
		book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		sheet = book.getSheetByName("Value");
		assertEquals("en_US", Locales.getCurrent().toString());

		//formula
		assertEquals(SCell.CellType.FORMULA, sheet.getCell(3,1).getType());
		assertEquals("SUM(10,20)", sheet.getCell(3,1).getFormulaValue());
		assertEquals("ISBLANK(B1)", sheet.getCell(3,2).getFormulaValue());
		assertEquals("B1", sheet.getCell(3,3).getFormulaValue());

		//Locale.GERMAN
		Locales.setThreadLocal(Locale.GERMANY);
		book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		sheet = book.getSheetByName("Value");
		assertEquals("de_DE", Locales.getCurrent().toString());

		//formula
		assertEquals(SCell.CellType.FORMULA, sheet.getCell(3,1).getType());
		assertEquals("SUM(10,20)", sheet.getCell(3,1).getFormulaValue());
		assertEquals("ISBLANK(B1)", sheet.getCell(3,2).getFormulaValue());
		assertEquals("B1", sheet.getCell(3,3).getFormulaValue());
	}
	/**
	 * Under different locales, TW and US, should import the same format pattern
	 */
	@Test
	public void formatNotDependLocaleTest(){
		SBook book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		SSheet sheet = book.getSheetByName("Format");
		assertEquals("zh_TW", Locales.getCurrent().toString());
		assertEquals("m/d/yyyy", sheet.getCell(1, 4).getCellStyle().getDataFormat());
		
		Locales.setThreadLocal(Locale.US);
		book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		sheet = book.getSheetByName("Format");
		assertEquals("en_US", Locales.getCurrent().toString());
		assertEquals("m/d/yyyy", sheet.getCell(1, 4).getCellStyle().getDataFormat());
	}
	
	@Test
	public void areaChart(){
		SBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		areaChart(book);
	}
	
	@Test
	public void barChart(){
		SBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		barChart(book);
		SSheet sheet = book.getSheetByName("Bar");
		SChart barChart = sheet.getChart(0);
		
		assertEquals("Bar Chart Title",barChart.getTitle());
		
	}
	
	@Test
	public void bubbleChart(){
		SBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		bubbleChart(book);
		SSheet sheet = book.getSheetByName("Bubble");
		SChart bubbleChart = sheet.getChart(0);
		assertEquals("Sales",bubbleChart.getTitle()); 
	}
	
	@Test
	public void columnChart(){
		SBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		columnChart(book);
	}
	
	@Test
	public void doughnutChart(){
		SBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		doughnutChart(book);
	}
	
	@Test
	public void lineChart(){
		SBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		lineChart(book);
		SSheet sheet = book.getSheetByName("Line");
		SChart line3dChart = sheet.getChart(1);
		assertEquals("Line 3D Title",line3dChart.getTitle()); 
	}
	
	@Test
	public void pieChart(){
		SBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		pieChart(book);
	}
	
	@Test
	public void scatterChart(){
		SBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		scatterChart(book);
	}
	
	@Ignore("XSSFStockChartData implementation is incorrect")
	public void stockChart(){
		SBook book = ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "Chart");
		SSheet sheet = book.getSheetByName("Stock");
		SChart stockChart = sheet.getChart(0);
		assertEquals(ChartType.STOCK, stockChart.getType());
		
		SGeneralChartData chartData = (SGeneralChartData)stockChart.getData();
		assertEquals(4, chartData.getNumOfSeries());
		assertEquals("Open", chartData.getSeries(0).getName());
		assertEquals("High", chartData.getSeries(1).getName());
		assertEquals("Low", chartData.getSeries(2).getName());
		assertEquals("Close", chartData.getSeries(3).getName());
	}
	
	/**
	 * quantity, width, height, format
	 */
	@Test
	public void picture(){
		SBook book = ImExpTestUtil.loadBook(PICTURE_IMPORT_FILE_UNDER_TEST, "XLSX");
		picture(book);
		
		SSheet sheet2 = book.getSheet(1);
		assertEquals(2,sheet2.getPictures().size());
		SPicture flowerJpg = sheet2.getPicture(0);
		assertEquals(Format.JPG, flowerJpg.getFormat());
		assertEquals(569, flowerJpg.getAnchor().getWidth());
		assertEquals(427, flowerJpg.getAnchor().getHeight());
		
		//different spec in XLS
		SPicture rainbowGif = sheet2.getPicture(1);
		assertEquals(Format.GIF, rainbowGif.getFormat());
		assertEquals(613, rainbowGif.getAnchor().getWidth());
		assertEquals(345, rainbowGif.getAnchor().getHeight());
	}
	
	@Test
	public void validation(){
		SBook book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XLSX");
		validation(book);
	}
	
	@Test
	public void autoFilter(){
		SBook book = ImExpTestUtil.loadBook(FILTER_IMPORT_FILE_UNDER_TEST, "XLSX");
		autoFilter(book);
	}
}



