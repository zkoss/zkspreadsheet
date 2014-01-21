package org.zkoss.zss.ngmodel;

import java.util.Locale;
import java.util.Set;

import junit.framework.Assert;

import org.junit.Before;
import org.junit.Test;
import org.zkoss.util.Locales;
import org.zkoss.zss.ngmodel.impl.AbstractBookSeriesAdv;
import org.zkoss.zss.ngmodel.impl.AbstractCellAdv;
import org.zkoss.zss.ngmodel.impl.RefImpl;
import org.zkoss.zss.ngmodel.impl.SheetImpl;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;

public class ModelFormulaTest {
	@Before
	public void beforeTest() {
		Locales.setThreadLocal(Locale.TAIWAN);
		SheetImpl.DEBUG = true;
	}
	
	protected NSheet initialDataGrid(NSheet sheet){
		return sheet;
	}
		

	@Test
	public void testFormulaDependencyClearAfterMoveCell(){
		NBook book = NBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		DependencyTable dt = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
		NSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("A1").setValue(3);
		sheet1.getCell("C3").setValue("=A1");
		
		Assert.assertEquals(3D, sheet1.getCell("C3").getValue());
		
		sheet1.getCell("A1").setValue(5);
		Assert.assertEquals(5D, sheet1.getCell("C3").getValue());
		
		
		Set<Ref> refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A1")));
		Assert.assertEquals(1, refs.size());
		Ref ref = refs.iterator().next();
		CellRegion refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("C3",refRegion.getReferenceString());
		
		//move cell
		sheet1.moveCell(new CellRegion("C3"), 1, 1);
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A1")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("D4",refRegion.getReferenceString());
		
		sheet1.getCell("A1").setValue(7);
		Assert.assertEquals(null, sheet1.getCell("C3").getValue());
		Assert.assertEquals(7D, sheet1.getCell("D4").getValue());
		
		sheet1.moveCell(new CellRegion("D4"), -2, -2);
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A1")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("B2",refRegion.getReferenceString());
		
		sheet1.getCell("A1").setValue(9);
		Assert.assertEquals(null, sheet1.getCell("D4").getValue());
		Assert.assertEquals(9D, sheet1.getCell("B2").getValue());
	}
	
//	@Test
	public void testFormulaDependencyClearAfterInsertDeleteCell(){
		//TODO
	}
	
	@Test
	public void testFormulaDependencyClearAfterInsertDeleteRow(){
		NBook book = NBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		DependencyTable dt = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
		NSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("A1").setValue(3);
		sheet1.getCell("C3").setValue("=A1");
		
		Assert.assertEquals(3D, sheet1.getCell("C3").getValue());
		
		sheet1.getCell("A1").setValue(5);
		Assert.assertEquals(5D, sheet1.getCell("C3").getValue());
		
		
		Set<Ref> refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A1")));
		Assert.assertEquals(1, refs.size());
		Ref ref = refs.iterator().next();
		CellRegion refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("C3",refRegion.getReferenceString());
		
		sheet1.insertRow(1, 1);
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A1")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("C4",refRegion.getReferenceString());
		
		sheet1.getCell("A1").setValue(7);
		Assert.assertEquals(null, sheet1.getCell("C3").getValue());
		Assert.assertEquals(7D, sheet1.getCell("C4").getValue());
		
		sheet1.deleteRow(1,2);
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A1")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("C2",refRegion.getReferenceString());
		
		sheet1.getCell("A1").setValue(9);
		Assert.assertEquals(null, sheet1.getCell("C4").getValue());
		Assert.assertEquals(9D, sheet1.getCell("C2").getValue());
	}
	
	@Test
	public void testFormulaDependencyClearAfterInsertDeleteColumn(){
		NBook book = NBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		DependencyTable dt = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
		NSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("A1").setValue(3);
		sheet1.getCell("C3").setValue("=A1");
		
		Assert.assertEquals(3D, sheet1.getCell("C3").getValue());
		
		sheet1.getCell("A1").setValue(5);
		Assert.assertEquals(5D, sheet1.getCell("C3").getValue());
		
		
		Set<Ref> refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A1")));
		Assert.assertEquals(1, refs.size());
		Ref ref = refs.iterator().next();
		CellRegion refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("C3",refRegion.getReferenceString());
		
		sheet1.insertColumn(1, 1);
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A1")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("D3",refRegion.getReferenceString());
		
		sheet1.getCell("A1").setValue(7);
		Assert.assertEquals(null, sheet1.getCell("C3").getValue());
		Assert.assertEquals(7D, sheet1.getCell("D3").getValue());
		
		sheet1.deleteColumn(1,2);
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A1")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("B3",refRegion.getReferenceString());
		
		sheet1.getCell("A1").setValue(9);
		Assert.assertEquals(null, sheet1.getCell("D3").getValue());
		Assert.assertEquals(9D, sheet1.getCell("B3").getValue());
	}
	
	//
	
	@Test
	public void testFormulaShiftingAfterMoveCell(){
		NBook book = NBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		DependencyTable dt = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
		NSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("A1").setValue(3);
		sheet1.getCell("C3").setValue("=A1");
		sheet1.getCell("G3").setValue("=C3");
		sheet1.getCell("A9").setValue("=SUM(C3)");
		sheet1.getCell("E6").setValue("=SUM(Sheet1!C3)");
		sheet1.getCell("F5").setValue("=SUM(Sheet1!C3)");
		
		Assert.assertEquals(3D, sheet1.getCell("C3").getValue());
		Assert.assertEquals(3D, sheet1.getCell("G3").getValue());
		Assert.assertEquals(3D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(3D, sheet1.getCell("E6").getValue());
		Assert.assertEquals(3D, sheet1.getCell("F5").getValue());
		
		sheet1.getCell("A1").setValue(5);
		Assert.assertEquals(5D, sheet1.getCell("C3").getValue());
		Assert.assertEquals(5D, sheet1.getCell("G3").getValue());
		Assert.assertEquals(5D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(5D, sheet1.getCell("E6").getValue());
		Assert.assertEquals(5D, sheet1.getCell("F5").getValue());
		
		//move cell
		sheet1.moveCell(new CellRegion("C3"), 1, 1);
		
		Assert.assertEquals("A1", sheet1.getCell("D4").getFormulaValue());
		Assert.assertEquals("D4", sheet1.getCell("G3").getFormulaValue());
		Assert.assertEquals("SUM(D4)", sheet1.getCell("A9").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!D4)", sheet1.getCell("E6").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!D4)", sheet1.getCell("F5").getFormulaValue());
		
		sheet1.getCell("A1").setValue(7);
		Assert.assertEquals(7D, sheet1.getCell("D4").getValue());
		Assert.assertEquals(7D, sheet1.getCell("G3").getValue());
		Assert.assertEquals(7D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(7D, sheet1.getCell("E6").getValue());
		Assert.assertEquals(7D, sheet1.getCell("F5").getValue());
		
		//move cell
		sheet1.moveCell(new CellRegion("D4"), -2, -1);
		
		Assert.assertEquals("C2", sheet1.getCell("G3").getFormulaValue());
		Assert.assertEquals("SUM(C2)", sheet1.getCell("A9").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!C2)", sheet1.getCell("E6").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!C2)", sheet1.getCell("F5").getFormulaValue());
		
		sheet1.getCell("A1").setValue(9);
		Assert.assertEquals(9D, sheet1.getCell("C2").getValue());
		Assert.assertEquals(9D, sheet1.getCell("G3").getValue());
		Assert.assertEquals(9D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(9D, sheet1.getCell("E6").getValue());
		Assert.assertEquals(9D, sheet1.getCell("F5").getValue());
	}
}
