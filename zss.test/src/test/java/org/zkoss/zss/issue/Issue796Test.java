package org.zkoss.zss.issue;

import java.io.File;
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
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.ErrorValue;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.impl.pdf.PdfExporter;

public class Issue796Test {
	
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
	 * Test parse table formula
	 */
	@Test
	public void tesParseTableFormula() {
		try {
			Book book = Util.loadBook(this, "book/796-table-formula.xlsx");
			Sheet sheet1 = book.getSheet("Sheet1");
			Range rngA3 = Ranges.range(sheet1, "A3");
			CellData a3 = rngA3.getCellData();
			Assert.assertEquals("Formula Type", CellData.CellType.FORMULA, a3.getType());
			//20141014, henrichen: We have not supported Table yet. So the result
			//is #Name!
//			Assert.assertEquals("Value a Number Type", CellData.CellType.NUMERIC, a3.getResultType());
//			Assert.assertEquals("Value", 1d, a3.getDoubleValue(), 0d);
			Assert.assertEquals("Formula", "SUBTOTAL(103,[Column1])", a3.getFormulaValue());
			Assert.assertEquals("Value an Error Type on 20141014", CellData.CellType.ERROR, a3.getResultType());
			Assert.assertEquals("Value", ErrorValue.NAME, a3.getValue());

			Range rngC2 = Ranges.range(sheet1, "C2");
			CellData c2 = rngC2.getCellData();
			Assert.assertEquals("Formula Type", CellData.CellType.FORMULA, c2.getType());
//			Assert.assertEquals("Value a Number Type", CellData.CellType.NUMERIC, c2.getResultType());
//			Assert.assertEquals("Value", 99d, c2.getDoubleValue(), 0d);
			Assert.assertEquals("Formula", "Table1[Column1]", c2.getFormulaValue());
			Assert.assertEquals("Value an Error Type on 20141014", CellData.CellType.ERROR, c2.getResultType());
			Assert.assertEquals("Value", ErrorValue.NAME, c2.getValue());
		} catch (Exception ex) {
			Assert.assertTrue("Parse Table Formula:" + ex.getMessage(), false);
		}
	}
}
