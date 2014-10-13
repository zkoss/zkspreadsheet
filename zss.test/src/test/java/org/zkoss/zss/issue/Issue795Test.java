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

public class Issue795Test {
	
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
	 * parse a defined name which contains question mark '?'
	 */
	@Test
	public void parsingQuestionMarkName() {
		try {
			Book book = Util.loadBook(this, "book/795-questionmark-name.xlsx");
			Sheet sheet1 = book.getSheet("Sheet1");
			Range rngA2 = Ranges.range(sheet1, "A2");
			CellData a2 = rngA2.getCellData();
			Assert.assertEquals("Formula Type", CellData.CellType.FORMULA, a2.getType());
			Assert.assertEquals("Value a Number Type", CellData.CellType.NUMERIC, a2.getResultType());
			Assert.assertEquals("Value", 100d, a2.getDoubleValue(), 0d);
		} catch (Exception ex) {
			ex.printStackTrace();
			Assert.assertTrue("parsingExternName Failed: " + ex.getMessage(), false);
		}
	}
}