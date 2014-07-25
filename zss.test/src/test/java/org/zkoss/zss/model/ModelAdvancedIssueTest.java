package org.zkoss.zss.model;

import java.lang.reflect.Method;
import java.util.Locale;
import java.util.Set;

import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.util.Locales;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.impl.AbstractBookSeriesAdv;
import org.zkoss.zss.model.sys.dependency.DependencyTable;
import org.zkoss.zss.model.sys.dependency.Ref;

public class ModelAdvancedIssueTest {

	@Before
	public void beforeTest() {
		Locales.setThreadLocal(Locale.TAIWAN);
	}
	
	protected SSheet initialDataGrid(SSheet sheet){
		return sheet;
	}

	@Test
	public void testModelCellDependencyAfterMove(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("A1").setValue(12);
		sheet1.getCell("B1").setValue(34);
		sheet1.getCell("C1").setValue("=SUM(A1:B1)");
		
		Assert.assertEquals(46D, sheet1.getCell("C1").getValue());
	
		DependencyTable table = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
		
		Set<Ref> refs = table.getDependents(getLastRef(sheet1.getCell("A1")));
		Assert.assertEquals(1, refs.size());
		Ref ref = refs.iterator().next();
		Assert.assertEquals("C1", new CellRegion(ref.getRow(),ref.getColumn()).getReferenceString());

		refs = table.getDependents(getLastRef(sheet1.getCell("B1")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		Assert.assertEquals("C1", new CellRegion(ref.getRow(),ref.getColumn()).getReferenceString());
		
		
		sheet1.getCell("C1").setValue("=SUM(A3:B3)");
		
		refs = table.getDependents(getLastRef(sheet1.getCell("A1")));
		Assert.assertEquals(0, refs.size());
		refs = table.getDependents(getLastRef(sheet1.getCell("B1")));
		Assert.assertEquals(0, refs.size());
		
		refs = table.getDependents(getLastRef(sheet1.getCell("A3")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		Assert.assertEquals("C1", new CellRegion(ref.getRow(),ref.getColumn()).getReferenceString());

		refs = table.getDependents(getLastRef(sheet1.getCell("B3")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		Assert.assertEquals("C1", new CellRegion(ref.getRow(),ref.getColumn()).getReferenceString());
		
		sheet1.getCell("A3").setValue(21);
		sheet1.getCell("B3").setValue(54);
		
		Assert.assertEquals(75D, sheet1.getCell("C1").getValue());
		
		
		//move
		sheet1.moveCell(new CellRegion("A3:B3"), 3, 0); //A3:B3 -> A6:B6
		
		Assert.assertEquals("SUM(A6:B6)", sheet1.getCell("C1").getFormulaValue());
		refs = table.getDependents(getLastRef(sheet1.getCell("A3")));
		Assert.assertEquals(0, refs.size());
		refs = table.getDependents(getLastRef(sheet1.getCell("B3")));
		Assert.assertEquals(0, refs.size());
		
		refs = table.getDependents(getLastRef(sheet1.getCell("A6")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		Assert.assertEquals("C1", new CellRegion(ref.getRow(),ref.getColumn()).getReferenceString());

		refs = table.getDependents(getLastRef(sheet1.getCell("B6")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		Assert.assertEquals("C1", new CellRegion(ref.getRow(),ref.getColumn()).getReferenceString());
		
		
		sheet1.getCell("A6").setValue(11);
		sheet1.getCell("B6").setValue(32);
		
		Assert.assertEquals(43D, sheet1.getCell("C1").getValue());
	
	}

	private Ref getLastRef(SCell cell) {
		try {
			Method method = cell.getClass().getDeclaredMethod("getRef", null);
			method.setAccessible(true);
			return (Ref)method.invoke(cell, null);
		} catch (Exception e) {
			throw new RuntimeException(e.getMessage(),e);
		}
	}
}
