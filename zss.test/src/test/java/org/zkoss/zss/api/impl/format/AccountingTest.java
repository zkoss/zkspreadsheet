package org.zkoss.zss.api.impl.format;

import static org.junit.Assert.assertEquals;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

public class AccountingTest {
	private static Book book;
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssLocale(Locale.TAIWAN);
		book = Util.loadBook("TestFile2007-Format.xlsx");
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssLocale();
	}
	
	@Test
	public void testAccounting3() {
		Sheet sheet = book.getSheet("Accounting");
		assertEquals(" NT$1,234,567,890.00 ", Ranges.range(sheet, "E3").getCellFormatText());
		assertEquals(" NT$1,234,567,890.00 ", Ranges.range(sheet, "G3").getCellFormatText());
	}
	
	@Test
	public void testAccounting5() {
		Sheet sheet = book.getSheet("Accounting");
		assertEquals(" NT$-1,234,567,890.00 ", Ranges.range(sheet, "E5").getCellFormatText());
		assertEquals(" NT$-1,234,567,890.00 ", Ranges.range(sheet, "G5").getCellFormatText());
	}
	
	@Test
	public void testAccounting7() {
		Sheet sheet = book.getSheet("Accounting");
		assertEquals(" $1,234,567,890.00 ", Ranges.range(sheet, "E7").getCellFormatText());
		assertEquals(" $1,234,567,890.00 ", Ranges.range(sheet, "G7").getCellFormatText());
	}
	
	@Test
	public void testAccounting9() {
		Sheet sheet = book.getSheet("Accounting");
		assertEquals("-$1,234,567,890.00 ", Ranges.range(sheet, "E9").getCellFormatText());
		assertEquals("-$1,234,567,890.00 ", Ranges.range(sheet, "G9").getCellFormatText());
	}
}
