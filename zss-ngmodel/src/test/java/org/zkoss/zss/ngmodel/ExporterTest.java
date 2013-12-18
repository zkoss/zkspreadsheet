package org.zkoss.zss.ngmodel;


import java.io.*;
import java.util.Locale;

import org.junit.*;
import org.zkoss.util.Locales;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;

/**
 * export test
 * @author kuro
 *
 */
public class ExporterTest extends ImExpTestBase {
	
	@BeforeClass
	static public void beforeClass() {
	}
	@Before
	public void beforeTest() {
		Locales.setThreadLocal(Locale.TAIWAN);
	}
	
	@Test
	public void sheetTest() {
		InputStream is = ImporterTest.class.getResourceAsStream("book/import.xlsx");
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(is));
		NBook book = ImExpTestUtil.loadBook(outFile);
		sheetTest(book);
	}
	
	@Test
	public void cellValueTest() {
		InputStream is = ImporterTest.class.getResourceAsStream("book/import.xlsx");
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(is));
		NBook book = ImExpTestUtil.loadBook(outFile);		
		cellValueTest(book);
	}

	@Test
	public void exportBorderFileTest() {
		InputStream is = ImporterTest.class.getResourceAsStream("book/cell_borders.xlsx");
		NBook book = ImExpTestUtil.loadBook(is);
		ImExpTestUtil.write(book);
	}

	@Test
	public void borderTypeTest() {
		// Use API to test
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
		
		File file = ImExpTestUtil.write(book);
		
		// FIXME assert it
		// confirm
		//cellBorderTest(inBook);
	}

	@Test
	public void exportXLSX() {
		InputStream is = ImporterTest.class.getResourceAsStream("book/import.xlsx");
		NBook book = ImExpTestUtil.loadBook(is);
		ImExpTestUtil.writeBookToFile(book, "./target/export.xlsx");
	}
	
	@Test
	public void exportXLS() {
		InputStream is = ImporterTest.class.getResourceAsStream("book/import.xls");
		NBook book = ImExpTestUtil.loadBook(is);
		ImExpTestUtil.writeBookToFile(book, "./target/export.xls");
	}
}
