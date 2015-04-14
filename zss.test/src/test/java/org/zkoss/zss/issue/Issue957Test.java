package org.zkoss.zss.issue;

import java.util.Locale;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

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
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.CellStyle.BorderType;
import org.zkoss.zss.api.model.EditableCellStyle;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.impl.CellImpl;
import org.zkoss.zss.range.SRange;
import org.zkoss.zss.range.SRanges;
import org.zkoss.zss.api.Exporters;

public class Issue957Test {
	
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
	 * Test if setCellStyle cause formula reevaluation.
	 */
	@Test
	public void testSetCellStyle() throws IOException {
		Book book = Util.loadBook(this, "book/blank.xlsx");
		Sheet sheet1 = book.getSheetAt(0);
		Range A1 = Ranges.range(sheet1, "A1");
		A1.setCellEditText("=A2");
		CellData value = A1.getCellData();
		Assert.assertEquals("A Formula", CellType.FORMULA, value.getType());
		Assert.assertEquals("CellValue", 0D, A1.getCellValue());
		
		
		SSheet sheet = sheet1.getInternalSheet();
		CellImpl cellA1 = (CellImpl) sheet.getCell("A1");
		Object valA1_0 = cellA1.getFromulaResultValue();
		Assert.assertNotNull("cached value", valA1_0);

		//change Style
		CellStyle oldStyle = A1.getCellStyle();
		EditableCellStyle newStyle = A1.getCellStyleHelper().createCellStyle(oldStyle);
		newStyle.setBorderBottom(BorderType.THIN);
		A1.setCellStyle(newStyle);
		Object valA1_4 = cellA1.getFromulaResultValue();
		Assert.assertNotNull("cached value after set style", valA1_4); //Formula Cache still there
		Assert.assertSame("FormulaResultValue", valA1_4, valA1_0);
		
		CellData value0 = A1.getCellData();
		Assert.assertEquals("A Formula", CellType.FORMULA, value0.getType());
		Assert.assertEquals("CellValue", 0D, A1.getCellValue());
		
		Object valA1_1 = cellA1.getFromulaResultValue();
		Assert.assertNotNull("cached value after set style and get value", valA1_1);
		Assert.assertSame("FormulaResultValue", valA1_1, valA1_0);
		
		//change precedence
		Range A2 = Ranges.range(sheet1, "A2");
		A2.setCellEditText("1");
		Object valA1_2 = cellA1.getFromulaResultValue();
		Assert.assertNull("cached value after set precedence", valA1_2); //Formula Cache is cleared
		
		CellData value1 = A1.getCellData();
		Assert.assertEquals("A Formula", CellType.FORMULA, value1.getType());
		Assert.assertEquals("CellValue", 1D, A1.getCellValue()); //get value to re-evaluate

		Object valA1_3 = cellA1.getFromulaResultValue();
		Assert.assertNotNull("cached value after set precedence and get value", valA1_3);
		Assert.assertNotSame("FormulaResultValue", valA1_3, valA1_0);
	}
}
