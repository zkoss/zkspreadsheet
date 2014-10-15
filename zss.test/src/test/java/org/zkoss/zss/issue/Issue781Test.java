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
import org.zkoss.zss.api.Range.InsertCopyOrigin;
import org.zkoss.zss.api.Range.InsertShift;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.api.model.Sheet;

/**
 * @author Hawk
 *
 */
public class Issue781Test {
	
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
	
	
	@Test
	public void testRenameSheet(){
		Book book;
		book = Util.loadBook(this, "book/781-rename-sheet.xlsx");
		Assert.assertEquals(Book.BookType.XLSX,book.getType());
		
		Sheet sheet2 = book.getSheet("Sheet2");
		Range rangeB1 = Ranges.range(sheet2, "B1");
		CellData data = rangeB1.getCellData();
		Assert.assertEquals("Sheet1!B1", data.getFormulaValue());
		
		Sheet sheet1 = book.getSheet("Sheet1");
		Assert.assertNotNull(sheet1);
		Range rangeSheet1 = Ranges.range(sheet1);
		rangeSheet1.setSheetName("St1");
		
		data = rangeB1.getCellData();
		Assert.assertEquals("St1!B1", data.getFormulaValue());
		
		Range rangeRow1 = Ranges.range(sheet1, "A1").toRowRange();
		rangeRow1.insert(InsertShift.DOWN, InsertCopyOrigin.FORMAT_NONE);

		data = rangeB1.getCellData();
		Assert.assertEquals("St1!B2", data.getFormulaValue());
	}
}
