package org.zkoss.zss.ngmodel;


import java.io.*;
import java.util.Locale;

import org.junit.*;
import org.zkoss.util.Locales;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;

/**
 * Common practice used in the test case:
 * 		1. We load and export the test file of importer then use importer's test cases to verify exported content.
 * 		2. Creating a book model in run-time to export and verify it.
 * @author kuro, Hawk
 *
 */
public class ExporterTest extends ImExpTestBase {
	
	@Before
	public void beforeTest() {
		Locales.setThreadLocal(Locale.TAIWAN);
	}
	
	@Test
	public void sheetTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook"));
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		sheetTest(book);
	}
	
	@Test
	public void cellValueTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook"));
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);		
		cellValueTest(book);
	}
	
	@Test
	public void sheetProtectionTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook"));
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		sheetProtectionTest(book);
	}
	@Test
	public void sheetNamedRangeTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook"));
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		sheetNamedRangeTest(book);
	}
	
	@Test
	public void cellStyleTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook"));
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		cellStyleTest(book);
	}
	
	@Test
	public void cellBorderTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook"));
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		cellBorderTest(book);
	}
	
	@Test
	public void cellFontNameTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook"));
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		cellFontNameTest(book);
	}
	
	@Test
	public void cellFontStyleTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook"));
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		cellFontStyleTest(book);
	}
	
	@Test
	public void cellFontColorTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook"));
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		cellFontColorTest(book);
	}
	
	@Test
	public void rowTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook"));
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		rowTest(book);
	}
	
	@Test
	public void cellFormatTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook"));
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		cellFormatTest(book);
	}
	
	@Test
	public void columnTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook"));
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		columnTest(book);
	}
	
	@Test
	public void lastChangedColumnTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook"));
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		lastChangedColumnTest(book);
	}
	
	@Test
	public void viewInfoTest() {
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook"));
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		viewInfoTest(book);
	}
	
	@Test
	public void mergedTest(){
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook"));
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		mergedTest(book);
	}

	// Use API to test
	@Test
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
	public void exportXLSX() {
		NBook book = ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook");
		ImExpTestUtil.writeBookToFile(book, ImExpTestUtil.DEFAULT_EXPORT_TARGET_PATH+"export.xlsx");
	}
	
	@Test
	public void exportXLS() {
		NBook book = ImExpTestUtil.loadBook(fileForImporterTest, "XSSFBook");
		ImExpTestUtil.writeBookToFile(book, ImExpTestUtil.DEFAULT_EXPORT_TARGET_PATH+"export.xls");
	}
}
