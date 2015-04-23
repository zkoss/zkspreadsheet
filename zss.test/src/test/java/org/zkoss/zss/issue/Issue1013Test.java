package org.zkoss.zss.issue;

import java.io.IOException;
import java.util.Locale;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

public class Issue1013Test {
	
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
	 * Read value from a cell which refer to an external Table column on another sheet.
	 * @throws IOException 
	 */
	@Test
	public void refTableOnAnotherSheet() throws IOException {
		Book book = Util.loadBook(this, "book/1013-ref-table.xlsx");
		Sheet sheet1 = book.getSheetAt(0);
		Range A1 = Ranges.range(sheet1, "A1");
		double value = A1.getCellData().getDoubleValue();
		Assert.assertEquals("Sheet1!A1", 10d, value, 0d);
	}
}