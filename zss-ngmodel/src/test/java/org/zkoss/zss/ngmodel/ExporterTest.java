package org.zkoss.zss.ngmodel;


import static org.junit.Assert.assertEquals;

import java.io.File;
import java.util.*;

import junit.framework.Assert;

import org.junit.*;
import org.zkoss.util.Locales;
import org.zkoss.zss.ngapi.impl.imexp.*;
import org.zkoss.zss.ngapi.impl.imexp.ExcelExportFactory.Type;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NFont.Boldweight;
import org.zkoss.zss.ngmodel.NFont.Underline;
import org.zkoss.zss.ngmodel.NPicture.Format;

/**
 * XLSX exporter test cases.
 * Common practices used in the test case:
 * 		1. We load and export the test file of importer then use importer's test cases to verify exported content.
 * 		2. Creating a book model in run-time to export and verify it.
 * @author kuro, Hawk
 *
 */
public class ExporterTest extends ImExpTestBase {
	
	protected Type EXPORTER_TYPE = ExcelExportFactory.Type.XLSX;

	@Before
	public void beforeTest() {
		Locales.setThreadLocal(Locale.TAIWAN);
	}
	
	@Test
	public void sheetTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		sheetTest(book);
	}
	
	@Test
	public void cellValueTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);		
		cellValueTest(book);
	}
	
	@Test
	public void sheetProtectionTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		sheetProtectionTest(book);
	}
	@Test
	public void sheetNamedRangeTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		sheetNamedRangeTest(book);
	}
	
	@Test
	public void cellStyleTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		cellStyleTest(book);
	}
	
	@Test
	public void cellBorderTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		cellBorderTest(book);
	}
	
	@Test
	public void cellFontNameTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		cellFontNameTest(book);
	}
	
	@Test
	public void cellFontStyleTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		cellFontStyleTest(book);
	}
	
	@Test
	public void cellFontColorTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		cellFontColorTest(book);
	}
	
	@Test
	public void rowTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		rowTest(book);
	}
	
	@Test
	public void cellFormatTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		cellFormatTest(book);
	}
	
	@Test
	public void columnTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		columnTest(book);
	}
	
	@Test
	public void lastChangedColumnTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		lastChangedColumnTest(book);
	}
	
	@Test
	public void viewInfoTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		viewInfoTest(book);
	}
	
	@Test
	public void mergedTest(){
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		mergedTest(book);
	}
	
	@Test
	public void richTextTest() {
		
	}
	
	@Test
	public void richTextModelTest() {
		NBook book = NBooks.createBook("rich");
		NSheet sheet = book.createSheet("first");
		NCell cell = sheet.getCell(0, 0);
		
		NRichText rText = cell.setupRichTextValue();
		NFont font1 = book.createFont(true);
		font1.setColor(book.createColor("#0000FF"));
		font1.setStrikeout(true);
		rText.addSegment("abc", font1);
		
		NFont font2 = book.createFont(true);
		font2.setColor(book.createColor("#FF0000"));
		font2.setBoldweight(Boldweight.BOLD);
		rText.addSegment("123", font2);
		
		NFont font3 = book.createFont(true);
		font3.setColor(book.createColor("#C78548"));
		font3.setUnderline(Underline.SINGLE);
		rText.addSegment("xyz", font3);
		
		cell = sheet.getCell(0, 1);
		rText = cell.setupRichTextValue();
		font1 = book.createFont(true);
		font1.setColor(book.createColor("#FFFF00"));
		font1.setItalic(true);
		rText.addSegment("Hello", font1);
		
		font2 = book.createFont(true);
		font2.setColor(book.createColor("#FF33FF"));
		font2.setBoldweight(Boldweight.BOLD);
		rText.addSegment("World", font2);
		
		font3 = book.createFont(true);
		font3.setColor(book.createColor("#CCCC99"));
		font3.setName("HGPSoeiKakupoptai");
		rText.addSegment("000", font3);
		
		ImExpTestUtil.write(book, ExcelExportFactory.Type.XLSX);
	}

	@Test @Ignore("incomplete")
	public void bookCreatedInRuntimeTest() {
		
		NBook book = NBooks.createBook("book1");
		
		NSheet sheet1 = book.createSheet("Sheet1");
		NCell cell1 = sheet1.getCell(1, 1);
		NCell cell2 = sheet1.getCell(1, 2);
		NCell cell3 = sheet1.getCell(1, 3);

		cell1.setStringValue("hair");
		cell2.setStringValue("dot");
		cell3.setStringValue("dash");

		NCellStyle style1 = book.createCellStyle(true);
		style1.setBorderBottom(BorderType.HAIR);
		cell1.setCellStyle(style1);

		NCellStyle style2 = book.createCellStyle(true);
		style2.setBorderBottom(BorderType.DOTTED);
		cell2.setCellStyle(style2);

		NCellStyle style3 = book.createCellStyle(true);
		style3.setBorderBottom(BorderType.DASHED);
		cell3.setCellStyle(style3);

		NCell cell21 = sheet1.getCell(2, 1);
		NCell cell22 = sheet1.getCell(2, 2);
		NCell cell23 = sheet1.getCell(2, 3);

		NCellStyle style21 = book.createCellStyle(true);
		style21.setBorderTop(BorderType.NONE);
		cell21.setCellStyle(style21);
		NCellStyle style22 = book.createCellStyle(true);
		style22.setBorderTop(BorderType.NONE);
		cell22.setCellStyle(style22);
		NCellStyle style23 = book.createCellStyle(true);
		style23.setBorderTop(BorderType.NONE);
		cell23.setCellStyle(style23);
		
		//File file = ImExpTestUtil.write(book);
		
		// FIXME assert it
		// confirm
		//cellBorderTest(inBook);
	}

	@Test
	public void export4HumanChecking() {
		NBook book = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		ImExpTestUtil.writeBookToFile(book, ImExpTestUtil.DEFAULT_EXPORT_TARGET_PATH+"humanChecking.xlsx", EXPORTER_TYPE);
	}
	
	
	@Test
	public void exportWidthSplitTest() {
		NBook book = NBooks.createBook("book1");
		NSheet sheet1 = book.createSheet("Sheet1");
		int defaultWidth = 100;
		sheet1.setDefaultColumnWidth(defaultWidth);
		sheet1.setDefaultRowHeight(200);
		Assert.assertEquals(defaultWidth, sheet1.getDefaultColumnWidth());
		Assert.assertEquals(200, sheet1.getDefaultRowHeight());
		
		Iterator<NColumnArray> arrays = sheet1.getColumnArrayIterator();
		Assert.assertFalse(arrays.hasNext());
		
		Assert.assertNull(sheet1.getColumnArray(0));
		
		sheet1.setupColumnArray(0, 8).setWidth(10);
		sheet1.setupColumnArray(11, 255);
		arrays = sheet1.getColumnArrayIterator();
		NColumnArray array = arrays.next();
		Assert.assertEquals(0, array.getIndex());
		Assert.assertEquals(8, array.getLastIndex());
		Assert.assertEquals(10, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(9, array.getIndex());
		Assert.assertEquals(10, array.getLastIndex());
		Assert.assertEquals(defaultWidth, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(11, array.getIndex());
		Assert.assertEquals(255, array.getLastIndex());
		Assert.assertEquals(defaultWidth, array.getWidth());
		
		///////////// first export
		File outFile = ImExpTestUtil.writeBookToFile(book, ImExpTestUtil.DEFAULT_EXPORT_TARGET_PATH+ImExpTestUtil.DEFAULT_EXPORT_FILE_NAME_XLSX, EXPORTER_TYPE);
		NBook outBook = ImExpTestUtil.loadBook(outFile, "OutBook");
		
		sheet1 = outBook.getSheet(0);
		
		// default width become 104px
		//Assert.assertEquals(defaultWidth, sheet1.getDefaultColumnWidth());
		Assert.assertEquals(200, sheet1.getDefaultRowHeight());

		arrays = sheet1.getColumnArrayIterator();
		Assert.assertTrue(arrays.hasNext());
		
		arrays = sheet1.getColumnArrayIterator();
		array = arrays.next();
		Assert.assertEquals(0, array.getIndex());
		Assert.assertEquals(8, array.getLastIndex());
		Assert.assertEquals(10, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(9, array.getIndex());
		Assert.assertEquals(10, array.getLastIndex());
		Assert.assertEquals(defaultWidth, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(11, array.getIndex());
		Assert.assertEquals(255, array.getLastIndex());
		Assert.assertEquals(defaultWidth, array.getWidth());
		
		///////////// second export
		File outFile2 = ImExpTestUtil.writeBookToFile(outBook, ImExpTestUtil.DEFAULT_EXPORT_TARGET_PATH+ImExpTestUtil.DEFAULT_EXPORT_FILE_NAME_XLSX, EXPORTER_TYPE);
		NBook outBook2 = ImExpTestUtil.loadBook(outFile2, "OutBook");
		
		sheet1 = outBook2.getSheet(0);
		
		// default width become 104px
		//Assert.assertEquals(100, sheet1.getDefaultColumnWidth());
		Assert.assertEquals(200, sheet1.getDefaultRowHeight());

		arrays = sheet1.getColumnArrayIterator();
		Assert.assertTrue(arrays.hasNext());
		
		arrays = sheet1.getColumnArrayIterator();
		array = arrays.next();
		Assert.assertEquals(0, array.getIndex());
		Assert.assertEquals(8, array.getLastIndex());
		Assert.assertEquals(10, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(9, array.getIndex());
		Assert.assertEquals(10, array.getLastIndex());
		Assert.assertEquals(100, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(11, array.getIndex());
		Assert.assertEquals(255, array.getLastIndex());
		Assert.assertEquals(100, array.getWidth());
	}
	
	@Test
	public void areaChart(){
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		areaChart(book);
	}
	
	@Test
	public void barChart(){
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		barChart(book);
	}
	
	@Test
	public void bubbleChart(){
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		bubbleChart(book);
	}
	
	@Test
	public void columnChart(){
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		columnChart(book);
	}
	
	@Test
	public void doughnutChart(){
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		doughnutChart(book);
	}
	
	@Test
	public void lineChart(){
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		lineChart(book);
	}
	
	@Test
	public void pieChart(){
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		pieChart(book);
	}
	
	@Test
	public void scatterChart(){
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(CHART_IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		scatterChart(book);
	}
	
	@Test
	public void picture(){
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(PICTURE_IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		picture(book);
		
		NSheet sheet2 = book.getSheet(1);
		assertEquals(2,sheet2.getPictures().size());
		NPicture flowerJpg = sheet2.getPicture(0);
		assertEquals(Format.JPG, flowerJpg.getFormat());
		assertEquals(569, flowerJpg.getAnchor().getWidth());
		assertEquals(427, flowerJpg.getAnchor().getHeight());
		
		//different spec in XLS
		NPicture rainbowGif = sheet2.getPicture(1);
		assertEquals(Format.GIF, rainbowGif.getFormat());
		assertEquals(613, rainbowGif.getAnchor().getWidth());
		assertEquals(345, rainbowGif.getAnchor().getHeight());
	}
	
	@Test
	public void validation(){
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		validation(book);
	}
	
	@Test
	public void autoFilter(){
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(FILTER_IMPORT_FILE_UNDER_TEST, "XSSFBook"), EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		autoFilter(book);
	}
}
