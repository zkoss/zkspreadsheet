package org.zkoss.zss.model;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import junit.framework.Assert;

import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.util.Locales;
import org.zkoss.zss.model.impl.AbstractBookSeriesAdv;
import org.zkoss.zss.model.impl.RefImpl;
import org.zkoss.zss.model.impl.SheetImpl;
import org.zkoss.zss.model.impl.sys.DependencyTableImpl;
import org.zkoss.zss.model.sys.dependency.DependencyTable;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.range.SRange;
import org.zkoss.zss.range.SRanges;
import org.zkoss.zss.range.SRange.DeleteShift;
import org.zkoss.zss.range.SRange.InsertCopyOrigin;
import org.zkoss.zss.range.SRange.InsertShift;

public class IssueTest {
	@BeforeClass
	static public void beforeClass() {
		Setup.touch();
	}
	@Before
	public void beforeTest() {
		Locales.setThreadLocal(Locale.TAIWAN);
	}
	
	@Test 
	public void testZSS581(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		
		SSheet sheet1 = book.createSheet("Sheet1");
		SSheet sheet2 = book.createSheet("Sheet2");
		SSheet sheet3 = book.createSheet("Sheet3");
		
		sheet1.getCell("A1").setValue("=Sheet2!A1");
		sheet2.getCell("A1").setValue("=Sheet3!A1");
		sheet3.getCell("A1").setValue("3");
		
		DependencyTable table = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
//		System.out.println(">>>>>>>>case1");
//		((DependencyTableImpl)table).dump();
		
		SRanges.range(sheet2).setSheetName("Two");
//		System.out.println(">>>>>>>>case2");
//		((DependencyTableImpl)table).dump();
		Assert.assertEquals("Two!A1",sheet1.getCell("A1").getFormulaValue());
		
		SRanges.range(sheet3).setSheetName("Three");
//		System.out.println(">>>>>>>>case3");
//		((DependencyTableImpl)table).dump();
		Assert.assertEquals("Three!A1",sheet2.getCell("A1").getFormulaValue());
		
			
		sheet3.getCell("A1").setValue(10);
		Assert.assertEquals(10D, sheet1.getCell("A1").getValue());
		Assert.assertEquals(10D, sheet2.getCell("A1").getValue());
		
	}
	
	@Test
	public void testZSS581_furthercheck(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		
		SSheet sheet1 = book.createSheet("Sheet1");
		SSheet sheet2 = book.createSheet("Sheet2");
		SSheet sheet3 = book.createSheet("Sheet3");
		
		sheet1.getCell("A1").setValue("=Sheet2!A1");
		sheet2.getCell("A1").setValue("=Sheet3!A1");
		sheet3.getCell("A1").setValue("3");
		
		DependencyTable table = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
//		((DependencyTableImpl)table).dump();
		
		List<Ref> refs = new ArrayList<Ref>(table.getDirectDependents(new RefImpl(book.getBookName())));
		Assert.assertEquals(2, refs.size());
		Ref ref = refs.get(0);
		Assert.assertEquals("Sheet1", ref.getSheetName());
		Assert.assertEquals(0, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(0, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());
		ref = refs.get(1);
		Assert.assertEquals("Sheet2", ref.getSheetName());
		Assert.assertEquals(0, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(0, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());		
		
		refs = new ArrayList<Ref>(table.getDirectDependents(new RefImpl(book.getBookName(),sheet3.getSheetName())));
		Assert.assertEquals(1, refs.size());
		ref = refs.get(0);
		Assert.assertEquals("Sheet2", ref.getSheetName());
		Assert.assertEquals(0, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(0, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());
		
		refs = new ArrayList<Ref>(table.getDirectDependents(new RefImpl(book.getBookName(),sheet2.getSheetName())));
		Assert.assertEquals(2, refs.size());
		ref = refs.get(0);
		Assert.assertEquals("Sheet1", ref.getSheetName());
		Assert.assertEquals(0, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(0, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());
		ref = refs.get(1);
		Assert.assertEquals("Sheet2", ref.getSheetName());
		Assert.assertEquals(0, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(0, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());
		
		refs = new ArrayList<Ref>(table.getDependents(new RefImpl(book.getBookName(),sheet3.getSheetName())));
		Assert.assertEquals(2, refs.size());
		ref = refs.get(0);
		Assert.assertEquals("Sheet2", ref.getSheetName());
		Assert.assertEquals(0, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(0, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());
		ref = refs.get(1);
		Assert.assertEquals("Sheet1", ref.getSheetName());
		Assert.assertEquals(0, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(0, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());
		
		refs = new ArrayList<Ref>(table.getDependents(new RefImpl(book.getBookName(),sheet2.getSheetName())));
		Assert.assertEquals(2, refs.size());
		ref = refs.get(0);
		Assert.assertEquals("Sheet1", ref.getSheetName());
		Assert.assertEquals(0, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(0, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());
		ref = refs.get(1);
		Assert.assertEquals("Sheet2", ref.getSheetName());
		Assert.assertEquals(0, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(0, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());		
	}	
	
	@Test
	public void testZSS582(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		
		SSheet sheet1 = book.createSheet("Sheet1");
		sheet1.getCell("A1").setValue(10D);
		sheet1.getCell("A2").setValue("=Sheet1!A1");
		Assert.assertEquals(10D, sheet1.getCell("A2").getValue());
		
		SRanges.range(sheet1).setSheetName(".ABC");
		Assert.assertEquals("'.ABC'!A1", sheet1.getCell("A2").getFormulaValue());
		
		sheet1.getCell("A1").setValue(20D);
		Assert.assertEquals(20D, sheet1.getCell("A2").getValue());
		
	}
	
	@Test 
	public void testZSS619_rowInsert(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = book.createSheet("Sheet1");
		
		sheet1.getCell("E4").setValue("=E4");
		Assert.assertEquals("#N/A", sheet1.getCell("E4").getErrorValue().getErrorString());
		
		sheet1.getCell("E4").setValue("=A1");
		
		Assert.assertEquals(0D, sheet1.getCell("E4").getValue());
		
		SRanges.range(sheet1,"4").insert(InsertShift.DEFAULT, InsertCopyOrigin.FORMAT_NONE);
		
		sheet1.getCell("A1").setValue(20);
		Assert.assertEquals("A1", sheet1.getCell("E5").getFormulaValue());
		Assert.assertEquals(20D, sheet1.getCell("E5").getValue());
		
		
		Assert.assertTrue(sheet1.getCell("E4").isNull());
		
		sheet1.getCell("E4").setValue("=E4");
		Assert.assertEquals("#N/A", sheet1.getCell("E4").getErrorValue().getErrorString());
	}
	
	@Test 
	public void testZSS619_rowDelete(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = book.createSheet("Sheet1");
		
		sheet1.getCell("E4").setValue("=E4");
		Assert.assertEquals("#N/A", sheet1.getCell("E4").getErrorValue().getErrorString());
		
		sheet1.getCell("E4").setValue("=A1");
		
		Assert.assertEquals(0D, sheet1.getCell("E4").getValue());
		
		SRanges.range(sheet1,"2").delete(DeleteShift.DEFAULT);
		
		sheet1.getCell("A1").setValue(20);
		Assert.assertEquals("A1", sheet1.getCell("E3").getFormulaValue());
		Assert.assertEquals(20D, sheet1.getCell("E3").getValue());
		
		
		Assert.assertTrue(sheet1.getCell("E4").isNull());
		
		sheet1.getCell("E4").setValue("=E4");
		Assert.assertEquals("#N/A", sheet1.getCell("E4").getErrorValue().getErrorString());
	}
	
	@Test 
	public void testZSS619_columnInsert(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = book.createSheet("Sheet1");
		
		sheet1.getCell("E4").setValue("=E4");
		Assert.assertEquals("#N/A", sheet1.getCell("E4").getErrorValue().getErrorString());
		
		sheet1.getCell("E4").setValue("=A1");
		
		Assert.assertEquals(0D, sheet1.getCell("E4").getValue());
		
		SRanges.range(sheet1,"E").insert(InsertShift.DEFAULT, InsertCopyOrigin.FORMAT_NONE);
		
		sheet1.getCell("A1").setValue(20);
		Assert.assertEquals("A1", sheet1.getCell("F4").getFormulaValue());
		Assert.assertEquals(20D, sheet1.getCell("F4").getValue());
		
		
		Assert.assertTrue(sheet1.getCell("E4").isNull());
		
		sheet1.getCell("E4").setValue("=E4");
		Assert.assertEquals("#N/A", sheet1.getCell("E4").getErrorValue().getErrorString());
	}
	
	@Test 
	public void testZSS619_columnDelete(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = book.createSheet("Sheet1");
		
		sheet1.getCell("E4").setValue("=E4");
		Assert.assertEquals("#N/A", sheet1.getCell("E4").getErrorValue().getErrorString());
		
		sheet1.getCell("E4").setValue("=A1");
		
		Assert.assertEquals(0D, sheet1.getCell("E4").getValue());
		
		SRanges.range(sheet1,"B").delete(DeleteShift.DEFAULT);
		
		sheet1.getCell("A1").setValue(20);
		Assert.assertEquals("A1", sheet1.getCell("D4").getFormulaValue());
		Assert.assertEquals(20D, sheet1.getCell("D4").getValue());
		
		
		Assert.assertTrue(sheet1.getCell("E4").isNull());
		
		sheet1.getCell("E4").setValue("=E4");
		Assert.assertEquals("#N/A", sheet1.getCell("E4").getErrorValue().getErrorString());
	}
	
	@Test 
	public void testZSS619_cellInsertV(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = book.createSheet("Sheet1");
		
		sheet1.getCell("E4").setValue("=E4");
		Assert.assertEquals("#N/A", sheet1.getCell("E4").getErrorValue().getErrorString());
		
		sheet1.getCell("E4").setValue("=A1");
		
		Assert.assertEquals(0D, sheet1.getCell("E4").getValue());
		
		SRanges.range(sheet1,"E4").insert(InsertShift.DOWN, InsertCopyOrigin.FORMAT_NONE);
		
		sheet1.getCell("A1").setValue(20);
		Assert.assertEquals("A1", sheet1.getCell("E5").getFormulaValue());
		Assert.assertEquals(20D, sheet1.getCell("E5").getValue());
		
		
		Assert.assertTrue(sheet1.getCell("E4").isNull());
		
		sheet1.getCell("E4").setValue("=E4");
		Assert.assertEquals("#N/A", sheet1.getCell("E4").getErrorValue().getErrorString());
	}
	
	@Test 
	public void testZSS619_cellDeleteV(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = book.createSheet("Sheet1");
		
		sheet1.getCell("E4").setValue("=E4");
		Assert.assertEquals("#N/A", sheet1.getCell("E4").getErrorValue().getErrorString());
		
		sheet1.getCell("E4").setValue("=A1");
		
		Assert.assertEquals(0D, sheet1.getCell("E4").getValue());
		
		SRanges.range(sheet1,"E2").delete(DeleteShift.UP);
		
		sheet1.getCell("A1").setValue(20);
		Assert.assertEquals("A1", sheet1.getCell("E3").getFormulaValue());
		Assert.assertEquals(20D, sheet1.getCell("E3").getValue());
		
		
		Assert.assertTrue(sheet1.getCell("E4").isNull());
		
		sheet1.getCell("E4").setValue("=E4");
		Assert.assertEquals("#N/A", sheet1.getCell("E4").getErrorValue().getErrorString());
	}
	
	@Test 
	public void testZSS619_cellInsertH(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = book.createSheet("Sheet1");
		
		sheet1.getCell("E4").setValue("=E4");
		Assert.assertEquals("#N/A", sheet1.getCell("E4").getErrorValue().getErrorString());
		
		sheet1.getCell("E4").setValue("=A1");
		
		Assert.assertEquals(0D, sheet1.getCell("E4").getValue());
		
		SRanges.range(sheet1,"E4").insert(InsertShift.RIGHT, InsertCopyOrigin.FORMAT_NONE);
		
		sheet1.getCell("A1").setValue(20);
		Assert.assertEquals("A1", sheet1.getCell("F4").getFormulaValue());
		Assert.assertEquals(20D, sheet1.getCell("F4").getValue());
		
		
		Assert.assertTrue(sheet1.getCell("E4").isNull());
		
		sheet1.getCell("E4").setValue("=E4");
		Assert.assertEquals("#N/A", sheet1.getCell("E4").getErrorValue().getErrorString());
	}
	
	@Test 
	public void testZSS619_cellDeleteH(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = book.createSheet("Sheet1");
		
		sheet1.getCell("E4").setValue("=E4");
		Assert.assertEquals("#N/A", sheet1.getCell("E4").getErrorValue().getErrorString());
		
		sheet1.getCell("E4").setValue("=A1");
		
		Assert.assertEquals(0D, sheet1.getCell("E4").getValue());
		
		SRanges.range(sheet1,"B4").delete(DeleteShift.LEFT);
		
		sheet1.getCell("A1").setValue(20);
		Assert.assertEquals("A1", sheet1.getCell("D4").getFormulaValue());
		Assert.assertEquals(20D, sheet1.getCell("D4").getValue());
		
		
		Assert.assertTrue(sheet1.getCell("E4").isNull());
		
		sheet1.getCell("E4").setValue("=E4");
		Assert.assertEquals("#N/A", sheet1.getCell("E4").getErrorValue().getErrorString());
	}
	
	
	@Test 
	public void testZSS626_cellDeleteH(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = book.createSheet("Sheet1");
		
		sheet1.getCell("C1").setValue(1);
		sheet1.getCell("C2").setValue(2);
		sheet1.getCell("C3").setValue(3);
		
		sheet1.getCell("A1").setValue("=SUM(C1:C3)");
		
		Assert.assertEquals(6D, sheet1.getCell("A1").getValue());
		
		SRanges.range(sheet1,"B1:B2").delete(DeleteShift.LEFT);
		
		Assert.assertEquals(1D, sheet1.getCell("B1").getValue());
		Assert.assertEquals(2D, sheet1.getCell("B2").getValue());
		
		Assert.assertEquals(null, sheet1.getCell("C1").getValue());
		Assert.assertEquals(null, sheet1.getCell("C2").getValue());
		Assert.assertEquals(3D, sheet1.getCell("C3").getValue());
		
		
		Assert.assertEquals(3D, sheet1.getCell("A1").getValue());
	}
	
	@Test 
	public void testZSS626_cellInsertH(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = book.createSheet("Sheet1");
		
		sheet1.getCell("C1").setValue(1);
		sheet1.getCell("C2").setValue(2);
		sheet1.getCell("C3").setValue(3);
		
		sheet1.getCell("A1").setValue("=SUM(C1:C3)");
		
		Assert.assertEquals(6D, sheet1.getCell("A1").getValue());
		
		SRanges.range(sheet1,"B1:B2").insert(InsertShift.RIGHT,InsertCopyOrigin.FORMAT_NONE);
		
		Assert.assertEquals(1D, sheet1.getCell("D1").getValue());
		Assert.assertEquals(2D, sheet1.getCell("D2").getValue());
		
		Assert.assertEquals(null, sheet1.getCell("C1").getValue());
		Assert.assertEquals(null, sheet1.getCell("C2").getValue());
		Assert.assertEquals(3D, sheet1.getCell("C3").getValue());
		
		
		Assert.assertEquals(3D, sheet1.getCell("A1").getValue());
	}
	
	@Test 
	public void testZSS652(){//test cut
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = book.createSheet("Sheet1");
		
		sheet1.getCell("A1").setValue(1);
		sheet1.getCell("B2").setValue(2);
		sheet1.getCell("C3").setValue("=A1");
		sheet1.getCell("D3").setValue("=B2");
		
		SRange range = SRanges.range(sheet1,"B2:D3");
		range.copy(SRanges.range(sheet1,"E2"),true);
		
		Assert.assertEquals("A1", sheet1.getCell("F3").getFormulaValue());
		Assert.assertEquals(1D, sheet1.getCell("F3").getValue());
		Assert.assertEquals("E2", sheet1.getCell("G3").getFormulaValue());
		Assert.assertEquals(2D, sheet1.getCell("G3").getValue());
		
		
		sheet1.getCell("A1").setValue(4);
		sheet1.getCell("E2").setValue(6);
		sheet1.getCell("B2").setValue(5);
		
		Assert.assertEquals(4D, sheet1.getCell("F3").getValue());
		Assert.assertEquals(6D, sheet1.getCell("G3").getValue());
	}
	
	@Test 
	public void testZSS652_1(){//test copy
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = book.createSheet("Sheet1");
		
		sheet1.getCell("A1").setValue(1);
		sheet1.getCell("B2").setValue(2);
		sheet1.getCell("C3").setValue("=A1");
		sheet1.getCell("D3").setValue("=B2");
		
		SRange range = SRanges.range(sheet1,"B2:D3");
		range.copy(SRanges.range(sheet1,"E2"),false);
		
		Assert.assertEquals("D1", sheet1.getCell("F3").getFormulaValue());
		Assert.assertEquals(0D, sheet1.getCell("F3").getValue());
		Assert.assertEquals("E2", sheet1.getCell("G3").getFormulaValue());
		Assert.assertEquals(2D, sheet1.getCell("G3").getValue());
		
		sheet1.getCell("D1").setValue(3);
		sheet1.getCell("A1").setValue(4);
		sheet1.getCell("E2").setValue(6);
		sheet1.getCell("B2").setValue(5);
		
		Assert.assertEquals(3D, sheet1.getCell("F3").getValue());
		Assert.assertEquals(6D, sheet1.getCell("G3").getValue());
	}	
	
	@Test 
	public void testZSS655(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = book.createSheet("Sheet1");
		
		book.createName("Test").setRefersToFormula("A1:B2");
		SRanges.range(sheet1,"A1").setValue(1);
		SRanges.range(sheet1,"B1").setValue(2);
		
		SRanges.range(sheet1,"C3").setEditText("=SUM(Test)");
		
		Assert.assertEquals(3D, sheet1.getCell("C3").getValue());
		
		SRanges.range(sheet1,"B1").setValue(3);
		Assert.assertEquals(4D, sheet1.getCell("C3").getValue());//fail on this line
		
	}
	
	@Test 
	public void testZSS655_further(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = book.createSheet("Sheet1");
		
		SRanges.range(sheet1,"C3").setEditText("=SUM(Test)");
		
		book.createName("Test").setRefersToFormula("A1:B2");
		SRanges.range(sheet1,"A1").setValue(1);
		SRanges.range(sheet1,"B1").setValue(2);
		
		
		Assert.assertEquals(3D, sheet1.getCell("C3").getValue());
		
		SRanges.range(sheet1,"B1").setValue(3);
		Assert.assertEquals(4D, sheet1.getCell("C3").getValue());//fail on this line
		
	}
}
