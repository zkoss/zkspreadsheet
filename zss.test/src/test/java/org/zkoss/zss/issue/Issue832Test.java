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

public class Issue832Test {
	
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
	public void checkHiddenSheets() {
		Book book = Util.loadBook(this, "book/832-hidden-sheets.xlsx");
		
		Sheet s1 = book.getSheet("Sheet1");
		Sheet s2 = book.getSheet("Sheet2");
		Sheet s3 = book.getSheet("Sheet3");
		Sheet s4 = book.getSheet("Sheet4");
		Sheet s5 = book.getSheet("Sheet5");
		Sheet s6 = book.getSheet("Sheet6");
		Sheet s7 = book.getSheet("Sheet7");
		
		Assert.assertEquals("sheet1 hidden", s1.isHidden(), true);
		Assert.assertEquals("sheet2 hidden", s2.isHidden(), false);
		Assert.assertEquals("sheet3 hidden", s3.isHidden(), true);
		Assert.assertEquals("sheet4 hidden", s4.isHidden(), false);
		Assert.assertEquals("sheet5 hidden", s5.isHidden(), true);
		Assert.assertEquals("sheet6 hidden", s6.isHidden(), false);
		Assert.assertEquals("sheet7 hidden", s7.isHidden(), false);

		Assert.assertEquals("sheet1 very hidden", s1.isVeryHidden(), false);
		Assert.assertEquals("sheet2 very hidden", s2.isVeryHidden(), false);
		Assert.assertEquals("sheet3 very hidden", s3.isVeryHidden(), false);
		Assert.assertEquals("sheet4 very hidden", s4.isVeryHidden(), false);
		Assert.assertEquals("sheet5 very hidden", s5.isVeryHidden(), false);
		Assert.assertEquals("sheet6 very hidden", s6.isVeryHidden(), false);
		Assert.assertEquals("sheet7 very hidden", s7.isVeryHidden(), false);

		Range rngSheet1 = Ranges.range(s1);
		rngSheet1.setSheetVisible(SheetVisible.VERY_HIDDEN);
		Assert.assertEquals("sheet1 hidden", s1.isHidden(), false);
		Assert.assertEquals("sheet1 very hidden", s1.isVeryHidden(), true);

		rngSheet1.setSheetVisible(SheetVisible.HIDDEN);
		Assert.assertEquals("sheet1 hidden", s1.isHidden(), true);
		Assert.assertEquals("sheet1 very hidden", s1.isVeryHidden(), false);

		rngSheet1.setSheetVisible(SheetVisible.VISIBLE);
		Assert.assertEquals("sheet1 hidden", s1.isHidden(), false);
		Assert.assertEquals("sheet1 very hidden", s1.isVeryHidden(), false);
	}
}
