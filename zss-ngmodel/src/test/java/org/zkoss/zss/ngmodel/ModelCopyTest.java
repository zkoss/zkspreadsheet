package org.zkoss.zss.ngmodel;

import java.util.Locale;

import junit.framework.Assert;

import org.junit.Before;
import org.junit.Test;
import org.zkoss.util.Locales;
import org.zkoss.zss.ngmodel.impl.SheetImpl;

public class ModelCopyTest {

	@Before
	public void beforeTest() {
		Locales.setThreadLocal(Locale.TAIWAN);
		SheetImpl.DEBUG = true;
	}
	
	protected NSheet initialDataGrid(NSheet sheet){
		return sheet;
	}
	
	@Test 
	public void testCopySimple(){
		NBook book = NBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		NSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("A1").setValue("A");
		sheet1.getCell("A2").setValue(13);
		sheet1.getCell("A3").setValue("=A2");
		sheet1.getCell("A4").setValue("=SUM(A2:A3)");
		
		sheet1.pasteCell(new SheetRegion(sheet1,"A1:A4"), new CellRegion("B2"), null);
		
		Assert.assertEquals("A", sheet1.getCell("B2").getValue());
		Assert.assertEquals(13D, sheet1.getCell("B3").getValue());
		Assert.assertEquals("B3", sheet1.getCell("B4").getFormulaValue());
		Assert.assertEquals("SUM(B3:B4)", sheet1.getCell("B5").getFormulaValue());
		
		Assert.assertEquals(13D, sheet1.getCell("B4").getValue());
		Assert.assertEquals(26D, sheet1.getCell("B5").getValue());
		
		sheet1.getCell("B3").setValue(10);
		Assert.assertEquals(10D, sheet1.getCell("B4").getValue());
		Assert.assertEquals(20D, sheet1.getCell("B5").getValue());
		
		//multiple row
		sheet1.pasteCell(new SheetRegion(sheet1,"A1:A4"), new CellRegion("C2:C9"), null);
		Assert.assertEquals("A", sheet1.getCell("C2").getValue());
		Assert.assertEquals(13D, sheet1.getCell("C3").getValue());
		Assert.assertEquals("C3", sheet1.getCell("C4").getFormulaValue());
		Assert.assertEquals("SUM(C3:C4)", sheet1.getCell("C5").getFormulaValue());
		Assert.assertEquals("A", sheet1.getCell("C6").getValue());
		Assert.assertEquals(13D, sheet1.getCell("C7").getValue());
		Assert.assertEquals("C7", sheet1.getCell("C8").getFormulaValue());
		Assert.assertEquals("SUM(C7:C8)", sheet1.getCell("C9").getFormulaValue());
		
		sheet1.getCell("C3").setValue(12);
		Assert.assertEquals(12D, sheet1.getCell("C4").getValue());
		Assert.assertEquals(24D, sheet1.getCell("C5").getValue());
		Assert.assertEquals(13D, sheet1.getCell("C8").getValue());
		Assert.assertEquals(26D, sheet1.getCell("C9").getValue());
		
		//multiple row/column
		sheet1.pasteCell(new SheetRegion(sheet1,"A1:A4"), new CellRegion("C2:D9"), null);
		Assert.assertEquals("A", sheet1.getCell("C2").getValue());
		Assert.assertEquals(13D, sheet1.getCell("C3").getValue());
		Assert.assertEquals("C3", sheet1.getCell("C4").getFormulaValue());
		Assert.assertEquals("SUM(C3:C4)", sheet1.getCell("C5").getFormulaValue());
		Assert.assertEquals("A", sheet1.getCell("C6").getValue());
		Assert.assertEquals(13D, sheet1.getCell("C7").getValue());
		Assert.assertEquals("C7", sheet1.getCell("C8").getFormulaValue());
		Assert.assertEquals("SUM(C7:C8)", sheet1.getCell("C9").getFormulaValue());
		//
		Assert.assertEquals("A", sheet1.getCell("D2").getValue());
		Assert.assertEquals(13D, sheet1.getCell("D3").getValue());
		Assert.assertEquals("D3", sheet1.getCell("D4").getFormulaValue());
		Assert.assertEquals("SUM(D3:D4)", sheet1.getCell("D5").getFormulaValue());
		Assert.assertEquals("A", sheet1.getCell("D6").getValue());
		Assert.assertEquals(13D, sheet1.getCell("D7").getValue());
		Assert.assertEquals("D7", sheet1.getCell("D8").getFormulaValue());
		Assert.assertEquals("SUM(D7:D8)", sheet1.getCell("D9").getFormulaValue());
		
		
		sheet1.getCell("C3").setValue(12);
		Assert.assertEquals(12D, sheet1.getCell("C4").getValue());
		Assert.assertEquals(24D, sheet1.getCell("C5").getValue());
		Assert.assertEquals(13D, sheet1.getCell("C8").getValue());
		Assert.assertEquals(26D, sheet1.getCell("C9").getValue());		
		
		sheet1.getCell("D7").setValue(12);
		Assert.assertEquals(13D, sheet1.getCell("D4").getValue());
		Assert.assertEquals(26D, sheet1.getCell("D5").getValue());
		Assert.assertEquals(12D, sheet1.getCell("D8").getValue());
		Assert.assertEquals(24D, sheet1.getCell("D9").getValue());
	}
	
	@Test 
	public void testCopySkipBlank(){
		NBook book = NBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		NSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("A1").setValue("A");
		sheet1.getCell("A3").setValue("");
		sheet1.getCell("A4").setValue("D");
		
		sheet1.getCell("B1").setValue("E");
		sheet1.getCell("B2").setValue("F");
		sheet1.getCell("B3").setValue("G");
		sheet1.getCell("B4").setValue("H");
		sheet1.getCell("C1").setValue("E");
		sheet1.getCell("C2").setValue("F");
		sheet1.getCell("C3").setValue("G");
		sheet1.getCell("C4").setValue("H");
		PasteOption opt = new PasteOption();
		
		sheet1.pasteCell(new SheetRegion(sheet1,"A1:A4"), new CellRegion("B1"), opt);
		
		Assert.assertEquals("A", sheet1.getCell("B1").getValue());
		Assert.assertEquals(null, sheet1.getCell("B2").getValue());
		Assert.assertEquals("", sheet1.getCell("B3").getValue());
		Assert.assertEquals("D", sheet1.getCell("B4").getValue());
		
		opt.setSkipBlank(true);
		sheet1.pasteCell(new SheetRegion(sheet1,"A1:A4"), new CellRegion("C1"), opt);
		Assert.assertEquals("A", sheet1.getCell("C1").getValue());
		Assert.assertEquals("F", sheet1.getCell("C2").getValue());
		Assert.assertEquals("", sheet1.getCell("C3").getValue());
		Assert.assertEquals("D", sheet1.getCell("C4").getValue());
	}
	
	@Test 
	public void testCut(){
		NBook book = NBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		NSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("A1").setValue("A");
		sheet1.getCell("A2").setValue("B");
		sheet1.getCell("A3").setValue("C");
		sheet1.getCell("A4").setValue("D");
		
		sheet1.getCell("B1").setValue("E");
		sheet1.getCell("B2").setValue("F");
		sheet1.getCell("B3").setValue("G");
		sheet1.getCell("B4").setValue("H");
		
		PasteOption opt = new PasteOption();
		opt.setCut(true);
		sheet1.pasteCell(new SheetRegion(sheet1,"A1:A4"), new CellRegion("B1"), opt);
		
		Assert.assertEquals(null, sheet1.getCell("A1").getValue());
		Assert.assertEquals(null, sheet1.getCell("A2").getValue());
		Assert.assertEquals(null, sheet1.getCell("A3").getValue());
		Assert.assertEquals(null, sheet1.getCell("A4").getValue());
		
		Assert.assertEquals("A", sheet1.getCell("B1").getValue());
		Assert.assertEquals("B", sheet1.getCell("B2").getValue());
		Assert.assertEquals("C", sheet1.getCell("B3").getValue());
		Assert.assertEquals("D", sheet1.getCell("B4").getValue());
		
		sheet1.pasteCell(new SheetRegion(sheet1,"B1:B4"), new CellRegion("B2"), opt);
		Assert.assertEquals(null, sheet1.getCell("B1").getValue());
		Assert.assertEquals("A", sheet1.getCell("B2").getValue());
		Assert.assertEquals("B", sheet1.getCell("B3").getValue());
		Assert.assertEquals("C", sheet1.getCell("B4").getValue());
		Assert.assertEquals("D", sheet1.getCell("B5").getValue());
		
		sheet1.pasteCell(new SheetRegion(sheet1,"B2:B4"), new CellRegion("B1"), opt);
		Assert.assertEquals("A", sheet1.getCell("B1").getValue());
		Assert.assertEquals("B", sheet1.getCell("B2").getValue());
		Assert.assertEquals("C", sheet1.getCell("B3").getValue());
		Assert.assertEquals(null, sheet1.getCell("B4").getValue());
		Assert.assertEquals("D", sheet1.getCell("B5").getValue());
	}
	

}
