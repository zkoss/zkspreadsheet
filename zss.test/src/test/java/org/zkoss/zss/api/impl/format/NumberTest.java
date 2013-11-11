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

public class NumberTest {
	
	private static Book book;
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
		book = Util.loadBook(NumberTest.class, "TestFile2007-Format.xlsx");
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssContextLocale(Locale.TAIWAN);
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssContextLocale();
	}
	
	@Test
	public void testNumberFormat3() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("123456789.00 ", Ranges.range(sheet, "E3").getCellFormatText());
		assertEquals("123456789.00 ", Ranges.range(sheet, "G3").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testNumberFormat5() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("(123456789.00)", Ranges.range(sheet, "E5").getCellFormatText());
		assertEquals("(123456789.00)", Ranges.range(sheet, "G5").getCellFormatText());
		// assertEquals("#ff0000", Ranges.range(sheet, "G5").getCellStyle().getFont().getColor().getHtmlColor());
	}
	
	@Test
	public void testNumberFormat7() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("123,456,789.00 ", Ranges.range(sheet, "E7").getCellFormatText());
		assertEquals("123,456,789.00 ", Ranges.range(sheet, "G7").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testNumberFormat9() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("(123,456,789.00)", Ranges.range(sheet, "E9").getCellFormatText());
		assertEquals("(123,456,789.00)", Ranges.range(sheet, "G9").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testNumberFormat11() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("123456789.00 ", Ranges.range(sheet, "E11").getCellFormatText());
		assertEquals("123456789.00 ", Ranges.range(sheet, "G11").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testNumberFormat13() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("(123456789.00)", Ranges.range(sheet, "E13").getCellFormatText());
		assertEquals("(123456789.00)", Ranges.range(sheet, "G13").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testNumberFormat15() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("123,456,789.00 ", Ranges.range(sheet, "E15").getCellFormatText());
		assertEquals("123,456,789.00 ", Ranges.range(sheet, "G15").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testNumberFormat17() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("(123,456,789.00)", Ranges.range(sheet, "E17").getCellFormatText());
		assertEquals("(123,456,789.00)", Ranges.range(sheet, "G17").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testNumberFormat19() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("123456789.00", Ranges.range(sheet, "E19").getCellFormatText());
		assertEquals("123456789.00", Ranges.range(sheet, "G19").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testNumberFormat21() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("123456789.00", Ranges.range(sheet, "E19").getCellFormatText());
		assertEquals("123456789.00", Ranges.range(sheet, "G19").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testNumberFormat23() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("123,456,789.00", Ranges.range(sheet, "E23").getCellFormatText());
		assertEquals("123,456,789.00", Ranges.range(sheet, "G23").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testNumberFormat25() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("-123,456,789.00 ", Ranges.range(sheet, "E25").getCellFormatText());
		assertEquals("-123,456,789.00 ", Ranges.range(sheet, "G25").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testNumberFormat27() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("123456789.00 ", Ranges.range(sheet, "E27").getCellFormatText());
		assertEquals("123456789.00 ", Ranges.range(sheet, "G27").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testNumberFormat29() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("-123456789.00 ", Ranges.range(sheet, "E29").getCellFormatText());
		assertEquals("-123456789.00 ", Ranges.range(sheet, "G29").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testNumberFormat31() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("123,456,789.00 ", Ranges.range(sheet, "E31").getCellFormatText());
		assertEquals("123,456,789.00 ", Ranges.range(sheet, "G31").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testNumberFormat33() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("-123,456,789.00 ", Ranges.range(sheet, "E33").getCellFormatText());
		assertEquals("-123,456,789.00 ", Ranges.range(sheet, "G33").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testNumberFormat35() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("123456789.00 ", Ranges.range(sheet, "E35").getCellFormatText());
		assertEquals("123456789.00 ", Ranges.range(sheet, "G35").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testNumberFormat37() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("-123456789.00 ", Ranges.range(sheet, "E37").getCellFormatText());
		assertEquals("-123456789.00 ", Ranges.range(sheet, "G37").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testNumberFormat39() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("123,456,789.00 ", Ranges.range(sheet, "E39").getCellFormatText());
		assertEquals("123,456,789.00 ", Ranges.range(sheet, "G39").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testNumberFormat41() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("-123,456,789.00 ", Ranges.range(sheet, "E41").getCellFormatText());
		assertEquals("-123,456,789.00 ", Ranges.range(sheet, "G41").getCellFormatText());
		// TODO confirm format color
	}
	
	@Test
	public void testNumberFormat43() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("1234.57 ", Ranges.range(sheet, "E43").getCellFormatText());
		assertEquals("1234.57 ", Ranges.range(sheet, "G43").getCellFormatText());
	}
	
	@Test
	public void testNumberFormat45() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("1234.568 ", Ranges.range(sheet, "E45").getCellFormatText());
		assertEquals("1234.568 ", Ranges.range(sheet, "G45").getCellFormatText());
	}

	@Test
	public void testNumberFormat47() {
		Sheet sheet = book.getSheet("Number");
		assertEquals("12.0", Ranges.range(sheet, "E47").getCellFormatText());
		assertEquals("12.0", Ranges.range(sheet, "G47").getCellFormatText());
	}
}
