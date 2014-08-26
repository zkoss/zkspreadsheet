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

public class SpecialTest {
	
	private static Book book;
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
		book = Util.loadBook(SpecialTest.class, "book/TestFile2007-Format.xlsx");
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
	public void testSpeical3() {
		Sheet sheet = book.getSheet("Special");
		assertEquals("804", Ranges.range(sheet, "E3").getCellFormatText());
		assertEquals("804", Ranges.range(sheet, "G3").getCellFormatText());
	}
	
	@Test
	public void testSpeical5() {
		Sheet sheet = book.getSheet("Special");
		assertEquals("15676988-4", Ranges.range(sheet, "E5").getCellFormatText());
		assertEquals("15676988-4", Ranges.range(sheet, "G5").getCellFormatText());
	}
	
	@Test
	public void testSpeical7() {
		Sheet sheet = book.getSheet("Special");
		assertEquals("(02) 123-4567", Ranges.range(sheet, "E7").getCellFormatText());
		assertEquals("(02) 123-4567", Ranges.range(sheet, "G7").getCellFormatText());
	}
	
	@Test
	public void testSpeical9() {
		Sheet sheet = book.getSheet("Special");
		assertEquals("(02) 1234-5678", Ranges.range(sheet, "E9").getCellFormatText());
		assertEquals("(02) 1234-5678", Ranges.range(sheet, "G9").getCellFormatText());
	}
	
	@Test
	public void testSpeical11() {
		Sheet sheet = book.getSheet("Special");
		assertEquals("0986-956-889", Ranges.range(sheet, "E11").getCellFormatText());
		assertEquals("0986-956-889", Ranges.range(sheet, "G11").getCellFormatText());
	}
}
