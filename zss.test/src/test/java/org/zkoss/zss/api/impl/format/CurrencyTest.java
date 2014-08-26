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

public class CurrencyTest {
	
	private static Book book;
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
		book = Util.loadBook(CurrencyTest.class, "book/TestFile2007-Format.xlsx");
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssLocale(Locale.TAIWAN);
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssLocale();
	}
	
	// ========================= Currency
	@Test
	public void testCurrency3() {
		Sheet sheet = book.getSheet("Currency");
		assertEquals("$1,234,567,890.00 ", Ranges.range(sheet, "E3").getCellFormatText());
		assertEquals("$1,234,567,890.00 ", Ranges.range(sheet, "G3").getCellFormatText());
	}
	
	@Test
	public void testCurrency5() {
		Sheet sheet = book.getSheet("Currency");
		assertEquals("($1,234,567,890.00)", Ranges.range(sheet, "E5").getCellFormatText());
		assertEquals("($1,234,567,890.00)", Ranges.range(sheet, "G5").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testCurrency7() {
		Sheet sheet = book.getSheet("Currency");
		assertEquals("$1,234,567,890.00 ", Ranges.range(sheet, "E7").getCellFormatText());
		assertEquals("$1,234,567,890.00 ", Ranges.range(sheet, "G7").getCellFormatText());
	}
	
	@Test
	public void testCurrency9() {
		Sheet sheet = book.getSheet("Currency");
		assertEquals("-$1,234,567,890.00", Ranges.range(sheet, "E9").getCellFormatText());
		assertEquals("-$1,234,567,890.00", Ranges.range(sheet, "G9").getCellFormatText());
	}
	
	@Test
	public void testCurrency11() {
		Sheet sheet = book.getSheet("Currency");
		assertEquals("$1,234,567,890.00", Ranges.range(sheet, "E11").getCellFormatText());
		assertEquals("$1,234,567,890.00", Ranges.range(sheet, "G11").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testCurrency13() {
		Sheet sheet = book.getSheet("Currency");
		assertEquals("$1,234,567,890.00", Ranges.range(sheet, "E13").getCellFormatText());
		assertEquals("$1,234,567,890.00", Ranges.range(sheet, "G13").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testCurrency15() {
		Sheet sheet = book.getSheet("Currency");
		assertEquals("$1,234,567,890.00", Ranges.range(sheet, "E15").getCellFormatText());
		assertEquals("$1,234,567,890.00", Ranges.range(sheet, "G15").getCellFormatText());
	}
	
	@Test
	public void testCurrency17() {
		Sheet sheet = book.getSheet("Currency");
		assertEquals("-$1,234,567,890.00", Ranges.range(sheet, "E17").getCellFormatText());
		assertEquals("-$1,234,567,890.00", Ranges.range(sheet, "G17").getCellFormatText());
	}
	
	@Test
	public void testCurrency19() {
		Sheet sheet = book.getSheet("Currency");
		assertEquals("$1,234,567,890.00", Ranges.range(sheet, "E19").getCellFormatText());
		assertEquals("$1,234,567,890.00", Ranges.range(sheet, "G19").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testCurrency21() {
		Sheet sheet = book.getSheet("Currency");
		assertEquals("-$1,234,567,890.00", Ranges.range(sheet, "E21").getCellFormatText());
		assertEquals("-$1,234,567,890.00", Ranges.range(sheet, "G21").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testCurrency23() {
		Sheet sheet = book.getSheet("Currency");
		assertEquals("NT$1,234,567,890.00 ", Ranges.range(sheet, "E23").getCellFormatText());
		assertEquals("NT$1,234,567,890.00 ", Ranges.range(sheet, "G23").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testCurrency25() {
		Sheet sheet = book.getSheet("Currency");
		assertEquals("(NT$1,234,567,890.00)", Ranges.range(sheet, "E25").getCellFormatText());
		assertEquals("(NT$1,234,567,890.00)", Ranges.range(sheet, "G25").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testCurrency27() {
		Sheet sheet = book.getSheet("Currency");
		assertEquals("US$1,234,567,890 ", Ranges.range(sheet, "E27").getCellFormatText());
		assertEquals("US$1,234,567,890 ", Ranges.range(sheet, "G27").getCellFormatText());
	}
	
	@Test
	public void testCurrency29() {
		Sheet sheet = book.getSheet("Currency");
		assertEquals("(US$1,234,567,890)", Ranges.range(sheet, "E29").getCellFormatText());
		assertEquals("(US$1,234,567,890)", Ranges.range(sheet, "G29").getCellFormatText());
	}
	
	@Test
	public void testCurrency31() {
		Sheet sheet = book.getSheet("Currency");
		assertEquals("US$1,234,567,890.00 ", Ranges.range(sheet, "E31").getCellFormatText());
		assertEquals("US$1,234,567,890.00 ", Ranges.range(sheet, "G31").getCellFormatText());
	}
	
	@Test
	public void testCurrency33() {
		Sheet sheet = book.getSheet("Currency");
		assertEquals("(US$1,234,567,890.00)", Ranges.range(sheet, "E33").getCellFormatText());
		assertEquals("(US$1,234,567,890.00)", Ranges.range(sheet, "G33").getCellFormatText());
	}
}
