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

public class FractionTest {
	
	private static Book book;
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
		book = Util.loadBook(FractionTest.class, "TestFile2007-Format.xlsx");
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
	public void testFraction3() {
		Sheet sheet = book.getSheet("Fraction");
		assertEquals(" 1/4", Ranges.range(sheet, "E3").getCellFormatText());
		assertEquals(" 1/4", Ranges.range(sheet, "G3").getCellFormatText());
	}
	
	@Test
	public void testFraction5() {
		Sheet sheet = book.getSheet("Fraction");
		assertEquals(" 21/25", Ranges.range(sheet, "E5").getCellFormatText());
		assertEquals(" 21/25", Ranges.range(sheet, "G5").getCellFormatText());
	}
	
	@Test
	public void testFraction7() {
		Sheet sheet = book.getSheet("Fraction");
		assertEquals(" 312/943", Ranges.range(sheet, "E7").getCellFormatText());
		assertEquals(" 312/943", Ranges.range(sheet, "G7").getCellFormatText());
	}
	
	@Test
	public void testFraction9() {
		Sheet sheet = book.getSheet("Fraction");
		assertEquals(" 1/2", Ranges.range(sheet, "E9").getCellFormatText());
		assertEquals(" 1/2", Ranges.range(sheet, "G9").getCellFormatText());
	}
	
	@Test
	public void testFraction11() {
		Sheet sheet = book.getSheet("Fraction");
		assertEquals(" 2/4", Ranges.range(sheet, "E11").getCellFormatText());
		assertEquals(" 2/4", Ranges.range(sheet, "G11").getCellFormatText());
	}
	
	@Test
	public void testFraction13() {
		Sheet sheet = book.getSheet("Fraction");
		assertEquals(" 4/8", Ranges.range(sheet, "E13").getCellFormatText());
		assertEquals(" 4/8", Ranges.range(sheet, "G13").getCellFormatText());
	}
	
	@Test
	public void testFraction15() {
		Sheet sheet = book.getSheet("Fraction");
		assertEquals("  8/16", Ranges.range(sheet, "E15").getCellFormatText());
		assertEquals("  8/16", Ranges.range(sheet, "G15").getCellFormatText());
	}
	
	@Test
	public void testFraction17() {
		Sheet sheet = book.getSheet("Fraction");
		assertEquals(" 3/10", Ranges.range(sheet, "E17").getCellFormatText());
		assertEquals(" 3/10", Ranges.range(sheet, "G17").getCellFormatText());
	}
	
	@Test
	public void testFraction19() {
		Sheet sheet = book.getSheet("Fraction");
		assertEquals(" 30/100", Ranges.range(sheet, "E19").getCellFormatText());
		assertEquals(" 30/100", Ranges.range(sheet, "G19").getCellFormatText());
	}
}
