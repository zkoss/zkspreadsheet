package org.zkoss.zss.issue;

import java.util.Locale;

import org.junit.Assert;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.SheetVisible;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

public class Issue845Test {
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssLocale(Locale.TAIWAN);
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssLocale();
	}

	/**
	 * Test render long text that across the cell boundary
	 */
	@Test
	public void checkIndirectNameRange() {
		Book book = Util.loadBook(this, "book/845-indirect-namerange.xlsx");
		
		Sheet s1 = book.getSheet("Sheet1");
		Range b2 = Ranges.range(s1, "B2");
		Assert.assertEquals("B2 original value", "15", b2.getCellFormatText());
		
		Range a1 = Ranges.range(s1, "A1");
		a1.setCellEditText("11");
		Assert.assertEquals("B2 value after change A1 to 11", "25", b2.getCellFormatText());
		
		Range b1 = Ranges.range(s1, "B1");
		b1.setCellEditText("item 2");
		Assert.assertEquals("B2 value after change B1 to \"item 2\"", "40", b2.getCellFormatText());

		Range a7 = Ranges.range(s1, "A7");
		a7.setCellEditText("16");
		Assert.assertEquals("B2 value after change A1 to 16", "50", b2.getCellFormatText());
	}
}
