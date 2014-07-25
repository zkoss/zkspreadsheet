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
public class CustomIgnoredTest {

	private static Book book;
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
		book = Util.loadBook("TestFile2007-Format.xlsx");
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
	public void testCustom3() { // Magenta
		Sheet sheet = book.getSheet("Custom");
		assertEquals("(123456789.00)", Ranges.range(sheet, "E3").getCellFormatText());
		assertEquals("#CA1F7B", Util.getFormatHTMLColor(sheet, 2, 4));
	}
	
	@Test
	public void testCustom11() { // CYAN
		Sheet sheet = book.getSheet("Custom");
		assertEquals("(123456789.00)", Ranges.range(sheet, "E11").getCellFormatText());
		assertEquals("#00ffff", Util.getFormatHTMLColor(sheet, 10, 4));
	}
}
