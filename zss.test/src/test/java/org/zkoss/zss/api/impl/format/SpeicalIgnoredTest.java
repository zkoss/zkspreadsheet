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
public class SpeicalIgnoredTest {
	
	private static Book book;
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
		book = Util.loadBook(SpecialTest.class, "TestFile2007-Format.xlsx");
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
	public void testSpeical13() {
		Sheet sheet = book.getSheet("Special");
		assertEquals("\u4E00\u842C\u4E8C\u5343\u4E09\u767E\u56DB\u5341\u4E94", Ranges.range(sheet, "E13").getCellFormatText());
		assertEquals("\u4E00\u842C\u4E8C\u5343\u4E09\u767E\u56DB\u5341\u4E94", Ranges.range(sheet, "G13").getCellFormatText());
	}
	
	@Test
	public void testSpeical15() {
		Sheet sheet = book.getSheet("Special");
		assertEquals("\u58F9\u842C\u8CB3\u4EDF\u53C3\u4F70\u8086\u62FE\u4F0D", Ranges.range(sheet, "E15").getCellFormatText());
		assertEquals("\u58F9\u842C\u8CB3\u4EDF\u53C3\u4F70\u8086\u62FE\u4F0D", Ranges.range(sheet, "G15").getCellFormatText());
	}
}
