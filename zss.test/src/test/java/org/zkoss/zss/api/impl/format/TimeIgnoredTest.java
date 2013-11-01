package org.zkoss.zss.api.impl.format;

import static org.junit.Assert.assertEquals;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

@Ignore
public class TimeIgnoredTest {
	
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
	public void testTime3() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("\u4E0A\u5348 11:10:50", Ranges.range(sheet, "E3").getCellFormatText());
		assertEquals("\u4E0A\u5348 11:10:50", Ranges.range(sheet, "G3").getCellFormatText());
	}
	
	@Test
	public void testTime7() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("11:10 AM", Ranges.range(sheet, "E7").getCellFormatText());
		assertEquals("11:10 AM", Ranges.range(sheet, "G7").getCellFormatText());
	}
	
	@Test
	public void testTime11() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("11:10:50 AM", Ranges.range(sheet, "E11").getCellFormatText());
		assertEquals("11:10:50 AM", Ranges.range(sheet, "G11").getCellFormatText());
	}
	
	@Test
	public void testTime13() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("1900/1/0 11:10 AM", Ranges.range(sheet, "E13").getCellFormatText());
		assertEquals("1900/1/0 11:10 AM", Ranges.range(sheet, "G13").getCellFormatText());
	}
	
	@Test
	public void testTime15() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("1900/1/0 11:10", Ranges.range(sheet, "E15").getCellFormatText());
		assertEquals("1900/1/0 11:10", Ranges.range(sheet, "G15").getCellFormatText());
	}
	
	@Test
	public void testTime21() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("\u4E0A\u534811\u664210\u5206", Ranges.range(sheet, "E21").getCellFormatText());
		assertEquals("\u4E0A\u534811\u664210\u5206", Ranges.range(sheet, "G21").getCellFormatText());
	}
	
	@Test
	public void testTime23() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("\u4E0A\u534811\u664210\u520650\u79D2", Ranges.range(sheet, "E23").getCellFormatText());
		assertEquals("\u4E0A\u534811\u664210\u520650\u79D2", Ranges.range(sheet, "G23").getCellFormatText());
	}
	
	@Test
	public void testTime33() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("1900/1/1  11:10:50 AM", Ranges.range(sheet, "E33").getCellFormatText());
		assertEquals("1900/1/1  11:10:50 AM", Ranges.range(sheet, "G33").getCellFormatText());
	}
	
	@Test
	public void testTime35() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("1900/1/2  11:10:50 AM", Ranges.range(sheet, "E35").getCellFormatText());
		assertEquals("1900/1/2  11:10:50 AM", Ranges.range(sheet, "G35").getCellFormatText());
	}
	
	
	@Test
	public void testTime37() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("11:10 AM", Ranges.range(sheet, "E37").getCellFormatText());
		assertEquals("11:10 AM", Ranges.range(sheet, "G37").getCellFormatText());
	}
	
	@Test
	public void testTime39() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("1900/1/2  11:10:50 AM", Ranges.range(sheet, "E39").getCellFormatText());
		assertEquals("1900/1/2  11:10:50 AM", Ranges.range(sheet, "G39").getCellFormatText());
	}
	
	@Test
	public void testTime41() {
		Sheet sheet = book.getSheet("Time");
		assertEquals("11:10:50 AM", Ranges.range(sheet, "E41").getCellFormatText());
		assertEquals("11:10:50 AM", Ranges.range(sheet, "G41").getCellFormatText());
	}	
	
}
