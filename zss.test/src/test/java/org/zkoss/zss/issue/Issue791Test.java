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
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.SCell.CellType;

public class Issue791Test {
	
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
	 * parse a formula which contains reference to a Name in another sheet.
	 */
	@Test
	public void nubmerInMergeCell() {
		try {
			Book book = Util.loadBook(this, "book/791-number-mergecell.xlsx");
			Assert.assertTrue(true);
			Sheet sheet1 = book.getSheet("Sheet1");
			Range rngA1 = Ranges.range(sheet1,  "A1");
			String a1 = rngA1.getCellFormatText();
			Assert.assertEquals("1.23E+16", a1);
			Range rngB1 = Ranges.range(sheet1, "B1");
			String b1 = rngB1.getCellFormatText();
			Assert.assertEquals("1.23457E+16", b1);
		} catch (Exception ex) {
			ex.printStackTrace();
			Assert.assertTrue("load Book Failed: " + ex.getMessage(), false);
		}
	}
}