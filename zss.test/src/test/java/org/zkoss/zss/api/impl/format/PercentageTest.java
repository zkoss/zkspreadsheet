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

public class PercentageTest {
	
	private static Book book;
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
		book = Util.loadBook(PercentageTest.class, "book/TestFile2007-Format.xlsx");
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssLocale(Locale.TAIWAN);
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssLocale();
	}
	
	// ========================= Percentage
	@Test
	public void testPercentage3() {
		Sheet sheet = book.getSheet("Percentage");
		assertEquals("98.59%", Ranges.range(sheet, "E3").getCellFormatText());
		assertEquals("98.59%", Ranges.range(sheet, "G3").getCellFormatText());
	}
	
	@Test
	public void testPercentage5() {
		Sheet sheet = book.getSheet("Percentage");
		assertEquals("98%", Ranges.range(sheet, "E5").getCellFormatText());
		assertEquals("98%", Ranges.range(sheet, "G5").getCellFormatText());
	}

}
