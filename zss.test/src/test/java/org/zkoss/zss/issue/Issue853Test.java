package org.zkoss.zss.issue;

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
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.api.model.CellData.CellType;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.range.SRange;
import org.zkoss.zss.range.SRanges;

public class Issue853Test {
	
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
	 * Test a cell with text format should see =xxx a string rather than a 
	 * formula
	 */
	@Test
	public void textFormat() {
		Book book = Util.loadBook(this, "book/853-string-format.xlsx");
		Sheet sheet1 = book.getSheetAt(0);
		Range b1 = Ranges.range(sheet1, "B1");
		Range a1 = Ranges.range(sheet1, "A1");
		a1.setCellEditText("=B1");
		b1.setCellEditText("1");
		Assert.assertEquals("=B1", "=B1", a1.getCellFormatText());
		CellData cd = a1.getCellData();
		CellType type = cd.getType();
		Assert.assertEquals("String type", CellType.STRING, type);
		Object value = cd.getValue();
		Assert.assertTrue("instanceof String", value instanceof String);
	}
}
