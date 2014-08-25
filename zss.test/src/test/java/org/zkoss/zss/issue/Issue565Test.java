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
import org.zkoss.zss.api.model.CellData.CellType;
import org.zkoss.zss.api.model.Sheet;

/**
 * Support input formula of different locale.
 * 
 * @author henrichen
 *
 */
public class Issue565Test {
	
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
	public void testInputFormula(){
		testInputFormulaTW(Util.loadBook(this, "book/blank.xlsx"));
		testInputFormulaTW(Util.loadBook(this, "book/blank.xls"));
		testInputFormulaDE(Util.loadBook(this, "book/blank.xlsx"));
		testInputFormulaDE(Util.loadBook(this, "book/blank.xls"));
	}
	
	public void testInputFormulaTW(Book book){
		Sheet sheet = book.getSheetAt(0);
		Range r;
		
		//TW
		Setup.setZssLocale(Locale.TAIWAN);
		
		//input
		r = Ranges.range(sheet,"A1");
		r.setCellEditText("=MAX(0,D$8-($A11+0.5))");
		
		Assert.assertEquals("General", r.getCellDataFormat()); //format change
		Assert.assertEquals("0", r.getCellFormatText());
		Assert.assertEquals(CellType.FORMULA, r.getCellData().getType());
		Assert.assertEquals("=MAX(0,D$8-($A11+0.5))", r.getCellEditText());
		Assert.assertEquals("MAX(0,D$8-($A11+0.5))", r.getCellData().getFormulaValue());
		
		//insert column at A
		r.toColumnRange().insert(InsertShift.RIGHT, InsertCopyOrigin.FORMAT_NONE);
		
		Range r2 = Ranges.range(sheet, "B1");
		Assert.assertEquals("General", r2.getCellDataFormat()); //format change
		Assert.assertEquals("0", r2.getCellFormatText());
		Assert.assertEquals(CellType.FORMULA, r2.getCellData().getType());
		Assert.assertEquals("=MAX(0,E$8-($B11+0.5))", r2.getCellEditText());
		Assert.assertEquals("MAX(0,E$8-($B11+0.5))", r2.getCellData().getFormulaValue());
		
		//copy/paste to B3
		Range r3 = Ranges.range(sheet, "B3");
		
		r2.paste(r3);
		Assert.assertEquals("General", r3.getCellDataFormat()); //format change
		Assert.assertEquals("0", r3.getCellFormatText());
		Assert.assertEquals(CellType.FORMULA, r3.getCellData().getType());
		Assert.assertEquals("=MAX(0,E$8-($B13+0.5))", r3.getCellEditText());
		Assert.assertEquals("MAX(0,E$8-($B13+0.5))", r3.getCellData().getFormulaValue());
	}
	
	public void testInputFormulaDE(Book book) {
		Sheet sheet = book.getSheetAt(0);
		Range r;
		
		//German
		Setup.setZssLocale(Locale.GERMANY);
		
		//input
		r = Ranges.range(sheet,"A1");
		r.setCellEditText("=MAX(0;D$8-($A11+0,5))");
		
		Assert.assertEquals("General", r.getCellDataFormat()); //format change
		Assert.assertEquals("0", r.getCellFormatText());
		Assert.assertEquals(CellType.FORMULA, r.getCellData().getType());
		Assert.assertEquals("=MAX(0;D$8-($A11+0,5))", r.getCellEditText());
		Assert.assertEquals("MAX(0,D$8-($A11+0.5))", r.getCellData().getFormulaValue());
		
		//insert column at A
		r.toColumnRange().insert(InsertShift.RIGHT, InsertCopyOrigin.FORMAT_NONE);
		
		Range r2 = Ranges.range(sheet, "B1");
		Assert.assertEquals("General", r2.getCellDataFormat()); //format change
		Assert.assertEquals("0", r2.getCellFormatText());
		Assert.assertEquals(CellType.FORMULA, r2.getCellData().getType());
		Assert.assertEquals("=MAX(0;E$8-($B11+0,5))", r2.getCellEditText());
		Assert.assertEquals("MAX(0,E$8-($B11+0.5))", r2.getCellData().getFormulaValue());

		//copy/paste to B3
		Range r3 = Ranges.range(sheet, "B3");
		
		r2.paste(r3);
		Assert.assertEquals("General", r3.getCellDataFormat()); //format change
		Assert.assertEquals("0", r3.getCellFormatText());
		Assert.assertEquals(CellType.FORMULA, r3.getCellData().getType());
		Assert.assertEquals("=MAX(0;E$8-($B13+0,5))", r3.getCellEditText());
		Assert.assertEquals("MAX(0,E$8-($B13+0.5))", r3.getCellData().getFormulaValue());
	}
}
