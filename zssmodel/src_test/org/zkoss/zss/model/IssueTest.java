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
import org.zkoss.zss.range.SRanges;

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
		((DependencyTableImpl)table).dump();
		
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
}
