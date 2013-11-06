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

public class TimeTest {
	
	private static Book book;
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
		book = Util.loadBook(TimeTest.class, "TestFile2007-Format.xlsx");
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
	public void testTime5() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("11:10", Ranges.range(sheet, "E5").getCellFormatText());
		assertEquals("11:10", Ranges.range(sheet, "G5").getCellFormatText());
	}
	
	@Test
	public void testTime9() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("11:10:50", Ranges.range(sheet, "E9").getCellFormatText());
		assertEquals("11:10:50", Ranges.range(sheet, "G9").getCellFormatText());
	}
	
	@Test
	public void testTime17() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("11\u664210\u5206", Ranges.range(sheet, "E17").getCellFormatText());
		assertEquals("11\u664210\u5206", Ranges.range(sheet, "G17").getCellFormatText());
	}
	
	@Test
	public void testTime19() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("11\u664210\u520650\u79D2", Ranges.range(sheet, "E19").getCellFormatText());
		assertEquals("11\u664210\u520650\u79D2", Ranges.range(sheet, "G19").getCellFormatText());
	}
	
	@Test
	public void testTime25() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("11:10:50", Ranges.range(sheet, "E25").getCellFormatText());
		assertEquals("11:10:50", Ranges.range(sheet, "G25").getCellFormatText());
	}
	
	@Test
	public void testTime27() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("10:50.0", Ranges.range(sheet, "E27").getCellFormatText());
		assertEquals("10:50.0", Ranges.range(sheet, "G27").getCellFormatText());
	}
	
	@Test
	public void testTime29() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("10:50", Ranges.range(sheet, "E29").getCellFormatText());
		assertEquals("10:50", Ranges.range(sheet, "G29").getCellFormatText());
	}
	
	@Test
	public void testTime31() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("11:10 \u4E0A\u5348", Ranges.range(sheet, "E31").getCellFormatText());
		assertEquals("11:10 \u4E0A\u5348", Ranges.range(sheet, "G31").getCellFormatText());
	}
	
	@Test
	public void testTime43() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("59", Ranges.range(sheet, "E43").getCellFormatText());
		assertEquals("59", Ranges.range(sheet, "G43").getCellFormatText());
	}
	
	@Test
	public void testTime45() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("11", Ranges.range(sheet, "E45").getCellFormatText());
		assertEquals("11", Ranges.range(sheet, "G45").getCellFormatText());
	}
	
	@Test
	public void testTime47() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("1", Ranges.range(sheet, "E47").getCellFormatText());
		assertEquals("1", Ranges.range(sheet, "G47").getCellFormatText());
	}
	
	@Test
	public void testTime49() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("3550", Ranges.range(sheet, "E49").getCellFormatText());
		assertEquals("3550", Ranges.range(sheet, "G49").getCellFormatText());
	}
	
	@Test
	public void testTime51() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("01", Ranges.range(sheet, "E51").getCellFormatText());
		assertEquals("01", Ranges.range(sheet, "G51").getCellFormatText());
	}
	
	@Test
	public void testTime53() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("50", Ranges.range(sheet, "E53").getCellFormatText());
		assertEquals("50", Ranges.range(sheet, "G53").getCellFormatText());
	}
	
	@Test
	public void testTime57() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("50", Ranges.range(sheet, "E57").getCellFormatText());
		assertEquals("50", Ranges.range(sheet, "G57").getCellFormatText());
	}
}
