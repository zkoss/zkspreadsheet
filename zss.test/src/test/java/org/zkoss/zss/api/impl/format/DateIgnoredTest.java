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
public class DateIgnoredTest {
	
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
	public void testDate5() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("2013/10/30", Ranges.range(sheet, "E5").getCellFormatText());
		assertEquals("2013/10/30", Ranges.range(sheet, "G5").getCellFormatText());
	}

	@Test
	public void testDate11() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("\u4E8C\u25CB\u4E00\u4E09\u5E74\u5341\u6708\u4E09\u5341\u65E5", Ranges.range(sheet, "E11").getCellFormatText());
		assertEquals("\u4E8C\u25CB\u4E00\u4E09\u5E74\u5341\u6708\u4E09\u5341\u65E5", Ranges.range(sheet, "G11").getCellFormatText());
	}
	
	
	@Test
	public void testDate13() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("\u5341\u6708\u4E09\u5341\u65E5", Ranges.range(sheet, "E13").getCellFormatText());
		assertEquals("\u5341\u6708\u4E09\u5341\u65E5", Ranges.range(sheet, "G13").getCellFormatText());
	}
	
	@Test
	public void testDate15() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("\u661F\u671F\u4E09", Ranges.range(sheet, "E15").getCellFormatText());
		assertEquals("\u661F\u671F\u4E09", Ranges.range(sheet, "G15").getCellFormatText());
	}
	
	@Test
	public void testDate17() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("\u9031\u4E09", Ranges.range(sheet, "E17").getCellFormatText());
		assertEquals("\u9031\u4E09", Ranges.range(sheet, "G17").getCellFormatText());
	}
	
	@Test
	public void testDate23() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("2013/10/30 12:00 AM", Ranges.range(sheet, "E23").getCellFormatText());
		assertEquals("2013/10/30 12:00 AM", Ranges.range(sheet, "G23").getCellFormatText());
	}
	
	@Test
	public void testDate31() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("30-Oct", Ranges.range(sheet, "E31").getCellFormatText());
		assertEquals("30-Oct", Ranges.range(sheet, "G31").getCellFormatText());
	}
	
	@Test
	public void testDate33() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("30-Oct-13", Ranges.range(sheet, "E33").getCellFormatText());
		assertEquals("30-Oct-13", Ranges.range(sheet, "G33").getCellFormatText());
	}
	
	@Test
	public void testDate35() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("O", Ranges.range(sheet, "E35").getCellFormatText());
		assertEquals("O", Ranges.range(sheet, "G35").getCellFormatText());
	}
	
	@Test
	public void testDate37() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("O-13", Ranges.range(sheet, "E37").getCellFormatText());
		assertEquals("O-13", Ranges.range(sheet, "G37").getCellFormatText());
	}
	
	@Test
	public void testDate39() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("102/10/30", Ranges.range(sheet, "E39").getCellFormatText());
		assertEquals("102/10/30", Ranges.range(sheet, "G39").getCellFormatText());
	}
	
	@Test
	public void testDate41() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("102\u5E7410\u670830\u65E5", Ranges.range(sheet, "E41").getCellFormatText());
		assertEquals("102\u5E7410\u670830\u65E5", Ranges.range(sheet, "G41").getCellFormatText());
	}
	
	@Test
	public void testDate43() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("102/10/30", Ranges.range(sheet, "E43").getCellFormatText());
		assertEquals("102/10/30", Ranges.range(sheet, "G43").getCellFormatText());
	}
	
	@Test
	public void testDate45() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("102\u5E7410\u670830\u65E5", Ranges.range(sheet, "E45").getCellFormatText());
		assertEquals("102\u5E7410\u670830\u65E5", Ranges.range(sheet, "G45").getCellFormatText());
	}
	
	@Test
	public void testDate47() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("30-Oct", Ranges.range(sheet, "E47").getCellFormatText());
		assertEquals("30-Oct", Ranges.range(sheet, "G47").getCellFormatText());
	}	
	
	@Test
	public void testDate49() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("30-Oct-13", Ranges.range(sheet, "E49").getCellFormatText());
		assertEquals("30-Oct-13", Ranges.range(sheet, "G49").getCellFormatText());
	}
	
	@Test
	public void testDate51() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("O", Ranges.range(sheet, "E51").getCellFormatText());
		assertEquals("O", Ranges.range(sheet, "G51").getCellFormatText());
	}
	
	@Test
	public void testDate53() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("O-13", Ranges.range(sheet, "E53").getCellFormatText());
		assertEquals("O-13", Ranges.range(sheet, "G53").getCellFormatText());
	}
	
	@Test
	public void testDate57() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("Jan", Ranges.range(sheet, "E57").getCellFormatText());
		assertEquals("Jan", Ranges.range(sheet, "G57").getCellFormatText());
	}
	
	@Test
	public void testDate59() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("January", Ranges.range(sheet, "E59").getCellFormatText());
		assertEquals("January", Ranges.range(sheet, "G59").getCellFormatText());
	}
	
	@Test
	public void testDate61() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("J", Ranges.range(sheet, "E61").getCellFormatText());
		assertEquals("J", Ranges.range(sheet, "G61").getCellFormatText());
	}
	
	@Test
	public void testDate67() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("Mon", Ranges.range(sheet, "E67").getCellFormatText());
		assertEquals("Mon", Ranges.range(sheet, "G67").getCellFormatText());
	}
	
	@Test
	public void testDate69() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("Monday", Ranges.range(sheet, "E69").getCellFormatText());
		assertEquals("Monday", Ranges.range(sheet, "G69").getCellFormatText());
	}

	@Test
	public void testDate73() {
		Sheet sheet = book.getSheet("Date");
		assertEquals("1990", Ranges.range(sheet, "E73").getCellFormatText());
		assertEquals("1990", Ranges.range(sheet, "G73").getCellFormatText());
	}	

}
