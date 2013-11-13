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

public class CustomTest {

	private static Book book;
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
		book = Util.loadBook(SpecialTest.class, "TestFile2007-Format.xlsx");
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
	public void testCustom19() {
		Sheet sheet = book.getSheet("Custom");
		assertEquals("$125.74 Surplus", Ranges.range(sheet, "E19").getCellFormatText());
		assertEquals("$125.74 Surplus", Ranges.range(sheet, "G19").getCellFormatText());
	}
	
	@Test
	public void testCustom21() {
		Sheet sheet = book.getSheet("Custom");
		assertEquals("$-125.74 Shortage", Ranges.range(sheet, "E21").getCellFormatText());
		assertEquals("$-125.74 Shortage", Ranges.range(sheet, "G21").getCellFormatText());
	}
	
	@Test
	public void testCustom23() {
		Sheet sheet = book.getSheet("Custom");
		assertEquals(">", Ranges.range(sheet, "E23").getCellFormatText());
		assertEquals(">", Ranges.range(sheet, "G23").getCellFormatText());
	}
	
	@Test
	public void testCustom25() {
		Sheet sheet = book.getSheet("Custom");
		assertEquals("~", Ranges.range(sheet, "E25").getCellFormatText());
		assertEquals("~", Ranges.range(sheet, "G25").getCellFormatText());
	}
	
	@Test
	public void testCustom27() {
		Sheet sheet = book.getSheet("Custom");
		assertEquals("<", Ranges.range(sheet, "E27").getCellFormatText());
		assertEquals("<", Ranges.range(sheet, "G27").getCellFormatText());
	}
	
	@Test
	public void testCustom29() {
		Sheet sheet = book.getSheet("Custom");
		assertEquals("!", Ranges.range(sheet, "E29").getCellFormatText());
		assertEquals("!", Ranges.range(sheet, "G29").getCellFormatText());
	}
	
	@Test
	public void testCustom31() {
		Sheet sheet = book.getSheet("Custom");
		assertEquals("5   1/  4", Ranges.range(sheet, "E31").getCellFormatText());
		assertEquals("5   1/  4", Ranges.range(sheet, "G31").getCellFormatText());
	}
	
	@Test
	public void testCustom33() {
		Sheet sheet = book.getSheet("Custom");
		assertEquals("  2.8  ", Ranges.range(sheet, "E33").getCellFormatText());
		assertEquals("  2.8  ", Ranges.range(sheet, "G33").getCellFormatText());
	}
	
	@Test
	public void testCustom35() {
		Sheet sheet = book.getSheet("Custom");
		assertEquals("8.900", Ranges.range(sheet, "E35").getCellFormatText());
		assertEquals("8.900", Ranges.range(sheet, "G35").getCellFormatText());
	}
}