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

/**
 * Tool for change other language to UTF-8
 * http://tool.chinaz.com/Tools/UTF-8.aspx
 * 
 * @author kuro
 *
 */
public class DateTest {
	
	private static Book book;
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
		book = Util.loadBook(DateTest.class, "TestFile2007-Format.xlsx");
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssLocale(Locale.TAIWAN);
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssLocale();
	}
	
	@Test
	public void testDate3() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("2013/10/30", Ranges.range(sheet, "E3").getCellFormatText());
		assertEquals("2013/10/30", Ranges.range(sheet, "G3").getCellFormatText());
	}
	
	@Test
	public void testDate7() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("2013\u5E7410\u670830\u65E5", Ranges.range(sheet, "E7").getCellFormatText());
		assertEquals("2013\u5E7410\u670830\u65E5", Ranges.range(sheet, "G7").getCellFormatText());
	}
	
	@Test
	public void testDate9() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("10\u670830\u65E5", Ranges.range(sheet, "E9").getCellFormatText());
		assertEquals("10\u670830\u65E5", Ranges.range(sheet, "G9").getCellFormatText());
	}
	
	@Test
	public void testDate19() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("2013/10/30", Ranges.range(sheet, "E19").getCellFormatText());
		assertEquals("2013/10/30", Ranges.range(sheet, "G19").getCellFormatText());
	}
	
	@Test
	public void testDate21() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("2013/10/30 0:00", Ranges.range(sheet, "E21").getCellFormatText());
		assertEquals("2013/10/30 0:00", Ranges.range(sheet, "G21").getCellFormatText());
	}
	
	@Test
	public void testDate25() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("10/30", Ranges.range(sheet, "E25").getCellFormatText());
		assertEquals("10/30", Ranges.range(sheet, "G25").getCellFormatText());
	}
	
	@Test
	public void testDate27() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("10/30/13", Ranges.range(sheet, "E27").getCellFormatText());
		assertEquals("10/30/13", Ranges.range(sheet, "G27").getCellFormatText());
	}
	
	@Test
	public void testDate29() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("10/30/13", Ranges.range(sheet, "E29").getCellFormatText());
		assertEquals("10/30/13", Ranges.range(sheet, "G29").getCellFormatText());
	}
	
	@Test
	public void testDate55() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("01", Ranges.range(sheet, "E55").getCellFormatText());
		assertEquals("01", Ranges.range(sheet, "G55").getCellFormatText());
	}
	
	@Test
	public void testDate63() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("2", Ranges.range(sheet, "E63").getCellFormatText());
		assertEquals("2", Ranges.range(sheet, "G63").getCellFormatText());
	}
	
	@Test
	public void testDate65() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("02", Ranges.range(sheet, "E65").getCellFormatText());
		assertEquals("02", Ranges.range(sheet, "G65").getCellFormatText());
	}
	
	@Test
	public void testDate71() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("00", Ranges.range(sheet, "E71").getCellFormatText());
		assertEquals("00", Ranges.range(sheet, "G71").getCellFormatText());
	}
}
