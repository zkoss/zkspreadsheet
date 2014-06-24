package org.zkoss.zss.model;

import java.util.Locale;
import java.util.Set;

import junit.framework.Assert;

import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.util.Locales;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SDataValidation;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SDataValidation.ValidationType;
import org.zkoss.zss.model.impl.AbstractBookSeriesAdv;
import org.zkoss.zss.model.impl.AbstractCellAdv;
import org.zkoss.zss.model.impl.RefImpl;
import org.zkoss.zss.model.impl.SheetImpl;
import org.zkoss.zss.model.sys.dependency.DependencyTable;
import org.zkoss.zss.model.sys.dependency.Ref;

public class ModelFormulaTest {
	@BeforeClass
	static public void beforeClass() {
		Setup.touch();
	}
	@Before
	public void beforeTest() {
		Locales.setThreadLocal(Locale.TAIWAN);
	}
	
	protected SSheet initialDataGrid(SSheet sheet){
		return sheet;
	}
		

	@Test
	public void testFormulaDependencyClearAfterMoveCell(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		DependencyTable dt = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
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
	
	@Test
	public void testFormulaDependencyClearAfterInsertDeleteCellVertical(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		DependencyTable dt = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
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
		
		sheet1.insertCell(new CellRegion("C3"),false);
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A1")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("C4",refRegion.getReferenceString());
		
		sheet1.getCell("A1").setValue(7);
		Assert.assertEquals(null, sheet1.getCell("C3").getValue());
		Assert.assertEquals(7D, sheet1.getCell("C4").getValue());
		
		sheet1.deleteCell(new CellRegion("C2:C3"),false);
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A1")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("C2",refRegion.getReferenceString());
		
		sheet1.getCell("A1").setValue(9);
		Assert.assertEquals(null, sheet1.getCell("C4").getValue());
		Assert.assertEquals(9D, sheet1.getCell("C2").getValue());
		
		
		//test trim
		sheet1.deleteCell(new CellRegion("C2:C4"),false);
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A1")));
		Assert.assertEquals(0, refs.size());

		
		sheet1.getCell("A1").setValue(9);
		Assert.assertEquals(null, sheet1.getCell("C4").getValue());
		Assert.assertEquals(null, sheet1.getCell("C2").getValue());
	}
	
	@Test
	public void testFormulaDependencyClearAfterInsertDeleteCellHorzontal(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		DependencyTable dt = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
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
		
		sheet1.insertCell(new CellRegion("C3"),true);
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A1")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("D3",refRegion.getReferenceString());
		
		sheet1.getCell("A1").setValue(7);
		Assert.assertEquals(null, sheet1.getCell("C3").getValue());
		Assert.assertEquals(7D, sheet1.getCell("D3").getValue());
		
		sheet1.deleteCell(new CellRegion("B3:C3"),true);
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A1")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("B3",refRegion.getReferenceString());
		
		sheet1.getCell("A1").setValue(9);
		Assert.assertEquals(null, sheet1.getCell("D3").getValue());
		Assert.assertEquals(9D, sheet1.getCell("B3").getValue());
		
		//test trim
		sheet1.deleteCell(new CellRegion("B3:D3"),true);
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A1")));
		Assert.assertEquals(0, refs.size());
		
		sheet1.getCell("A1").setValue(9);
		Assert.assertEquals(null, sheet1.getCell("D3").getValue());
		Assert.assertEquals(null, sheet1.getCell("B3").getValue());
		
	}
	
	@Test
	public void testFormulaDependencyClearAfterInsertDeleteRow(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		DependencyTable dt = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
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
		
		sheet1.deleteRow(1,1);
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A1")));
		Assert.assertEquals(0, refs.size());
		
		sheet1.getCell("A1").setValue(9);
		Assert.assertEquals(null, sheet1.getCell("C4").getValue());
		Assert.assertEquals(null, sheet1.getCell("C2").getValue());
	}
	
	@Test
	public void testFormulaDependencyClearAfterInsertDeleteColumn(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		DependencyTable dt = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
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
		
		sheet1.deleteColumn(1,1);
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A1")));
		Assert.assertEquals(0, refs.size());
		
		sheet1.getCell("A1").setValue(9);
		Assert.assertEquals(null, sheet1.getCell("D3").getValue());
		Assert.assertEquals(null, sheet1.getCell("B3").getValue());
	}
	
	//
	
	@Test
	public void testFormulaShiftAfterMoveCell(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("A1").setValue(3);
		sheet1.getCell("C3").setValue("=A1");
		sheet1.getCell("G1").setValue("=C3");
		sheet1.getCell("A9").setValue("=SUM(C3)");
		sheet1.getCell("E6").setValue("=SUM(Sheet1!C3)");
		sheet1.getCell("F5").setValue("=SUM(Sheet1!C3)");
		
		Assert.assertEquals(3D, sheet1.getCell("C3").getValue());
		Assert.assertEquals(3D, sheet1.getCell("G1").getValue());
		Assert.assertEquals(3D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(3D, sheet1.getCell("E6").getValue());
		Assert.assertEquals(3D, sheet1.getCell("F5").getValue());
		
		sheet1.getCell("A1").setValue(5);
		Assert.assertEquals(5D, sheet1.getCell("C3").getValue());
		Assert.assertEquals(5D, sheet1.getCell("G1").getValue());
		Assert.assertEquals(5D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(5D, sheet1.getCell("E6").getValue());
		Assert.assertEquals(5D, sheet1.getCell("F5").getValue());
		
		//move cell
		sheet1.moveCell(new CellRegion("C3"), 1, 1);
		
		Assert.assertEquals("A1", sheet1.getCell("D4").getFormulaValue());
		Assert.assertEquals("D4", sheet1.getCell("G1").getFormulaValue());
		Assert.assertEquals("SUM(D4)", sheet1.getCell("A9").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!D4)", sheet1.getCell("E6").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!D4)", sheet1.getCell("F5").getFormulaValue());
		
		sheet1.getCell("A1").setValue(7);
		Assert.assertEquals(7D, sheet1.getCell("D4").getValue());
		Assert.assertEquals(7D, sheet1.getCell("G1").getValue());
		Assert.assertEquals(7D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(7D, sheet1.getCell("E6").getValue());
		Assert.assertEquals(7D, sheet1.getCell("F5").getValue());
		
		//move cell
		sheet1.moveCell(new CellRegion("D4"), -2, -1);
		
		Assert.assertEquals("C2", sheet1.getCell("G1").getFormulaValue());
		Assert.assertEquals("SUM(C2)", sheet1.getCell("A9").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!C2)", sheet1.getCell("E6").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!C2)", sheet1.getCell("F5").getFormulaValue());
		
		sheet1.getCell("A1").setValue(9);
		Assert.assertEquals(9D, sheet1.getCell("C2").getValue());
		Assert.assertEquals(9D, sheet1.getCell("G1").getValue());
		Assert.assertEquals(9D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(9D, sheet1.getCell("E6").getValue());
		Assert.assertEquals(9D, sheet1.getCell("F5").getValue());
	}
	
	@Test
	public void testFormulaShiftAfterInsertDeleteCellVertical(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("A1").setValue(3);
		sheet1.getCell("C3").setValue("=A1");
		sheet1.getCell("G1").setValue("=C3");
		sheet1.getCell("A9").setValue("=SUM(C3)");
		sheet1.getCell("E6").setValue("=SUM(Sheet1!C3)");
		sheet1.getCell("F5").setValue("=SUM(Sheet1!C3,1)");
		
		Assert.assertEquals(3D, sheet1.getCell("C3").getValue());
		Assert.assertEquals(3D, sheet1.getCell("G1").getValue());
		Assert.assertEquals(3D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(3D, sheet1.getCell("E6").getValue());
		Assert.assertEquals(4D, sheet1.getCell("F5").getValue());
		
		sheet1.getCell("A1").setValue(5);
		Assert.assertEquals(5D, sheet1.getCell("C3").getValue());
		Assert.assertEquals(5D, sheet1.getCell("G1").getValue());
		Assert.assertEquals(5D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(5D, sheet1.getCell("E6").getValue());
		Assert.assertEquals(6D, sheet1.getCell("F5").getValue());
		
		sheet1.insertCell(new CellRegion("C2:E2"),false);
		
		Assert.assertEquals("A1", sheet1.getCell("C4").getFormulaValue());
		Assert.assertEquals("C4", sheet1.getCell("G1").getFormulaValue());
		Assert.assertEquals("SUM(C4)", sheet1.getCell("A9").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!C4)", sheet1.getCell("E7").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!C4,1)", sheet1.getCell("F5").getFormulaValue());
		
		sheet1.getCell("A1").setValue(7);
		Assert.assertEquals(7D, sheet1.getCell("C4").getValue());
		Assert.assertEquals(7D, sheet1.getCell("G1").getValue());
		Assert.assertEquals(7D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(7D, sheet1.getCell("E7").getValue());
		Assert.assertEquals(8D, sheet1.getCell("F5").getValue());
		
		//move cell
		sheet1.deleteCell(new CellRegion("C2:E3"),false);
		Assert.assertEquals("A1", sheet1.getCell("C2").getFormulaValue());
		Assert.assertEquals("C2", sheet1.getCell("G1").getFormulaValue());
		Assert.assertEquals("SUM(C2)", sheet1.getCell("A9").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!C2)", sheet1.getCell("E5").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!C2,1)", sheet1.getCell("F5").getFormulaValue());
		
		sheet1.getCell("A1").setValue(9);
		Assert.assertEquals(9D, sheet1.getCell("C2").getValue());
		Assert.assertEquals(9D, sheet1.getCell("G1").getValue());
		Assert.assertEquals(9D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(9D, sheet1.getCell("E5").getValue());
		Assert.assertEquals(10D, sheet1.getCell("F5").getValue());
		
		//test trim
		sheet1.deleteCell(new CellRegion("C2"),false);
		Assert.assertEquals("#REF!", sheet1.getCell("G1").getFormulaValue());
		Assert.assertEquals("SUM(#REF!)", sheet1.getCell("A9").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!#REF!)", sheet1.getCell("E5").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!#REF!,1)", sheet1.getCell("F5").getFormulaValue());
	}
	
	
	@Test
	public void testFormulaShiftAfterInsertDeleteCellHorzontal(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("A1").setValue(3);
		sheet1.getCell("C3").setValue("=A1");
		sheet1.getCell("G1").setValue("=C3");
		sheet1.getCell("A9").setValue("=SUM(C3)");
		sheet1.getCell("E6").setValue("=SUM(Sheet1!C3)");
		sheet1.getCell("F5").setValue("=SUM(Sheet1!C3,1)");
		
		Assert.assertEquals(3D, sheet1.getCell("C3").getValue());
		Assert.assertEquals(3D, sheet1.getCell("G1").getValue());
		Assert.assertEquals(3D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(3D, sheet1.getCell("E6").getValue());
		Assert.assertEquals(4D, sheet1.getCell("F5").getValue());
		
		sheet1.getCell("A1").setValue(5);
		Assert.assertEquals(5D, sheet1.getCell("C3").getValue());
		Assert.assertEquals(5D, sheet1.getCell("G1").getValue());
		Assert.assertEquals(5D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(5D, sheet1.getCell("E6").getValue());
		Assert.assertEquals(6D, sheet1.getCell("F5").getValue());
		
		sheet1.insertCell(new CellRegion("B3:B5"),true);
		
		Assert.assertEquals("A1", sheet1.getCell("D3").getFormulaValue());
		Assert.assertEquals("D3", sheet1.getCell("G1").getFormulaValue());
		Assert.assertEquals("SUM(D3)", sheet1.getCell("A9").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!D3)", sheet1.getCell("E6").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!D3,1)", sheet1.getCell("G5").getFormulaValue());
		
		sheet1.getCell("A1").setValue(7);
		Assert.assertEquals(7D, sheet1.getCell("D3").getValue());
		Assert.assertEquals(7D, sheet1.getCell("G1").getValue());
		Assert.assertEquals(7D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(7D, sheet1.getCell("E6").getValue());
		Assert.assertEquals(8D, sheet1.getCell("G5").getValue());
		

		sheet1.deleteCell(new CellRegion("B3:C5"),true);
		Assert.assertEquals("A1", sheet1.getCell("B3").getFormulaValue());
		Assert.assertEquals("B3", sheet1.getCell("G1").getFormulaValue());
		Assert.assertEquals("SUM(B3)", sheet1.getCell("A9").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!B3)", sheet1.getCell("E6").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!B3,1)", sheet1.getCell("E5").getFormulaValue());
		
		sheet1.getCell("A1").setValue(9);
		Assert.assertEquals(9D, sheet1.getCell("B3").getValue());
		Assert.assertEquals(9D, sheet1.getCell("G1").getValue());
		Assert.assertEquals(9D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(9D, sheet1.getCell("E6").getValue());
		Assert.assertEquals(10D, sheet1.getCell("E5").getValue());
		
		//trim
		sheet1.deleteCell(new CellRegion("B3"),true);
		Assert.assertEquals("#REF!", sheet1.getCell("G1").getFormulaValue());
		Assert.assertEquals("SUM(#REF!)", sheet1.getCell("A9").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!#REF!)", sheet1.getCell("E6").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!#REF!,1)", sheet1.getCell("E5").getFormulaValue());
	}
	
	@Test
	public void testFormulaShiftAfterInsertDeleteRow(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("A1").setValue(3);
		sheet1.getCell("C3").setValue("=A1");
		sheet1.getCell("G1").setValue("=C3");
		sheet1.getCell("A9").setValue("=SUM(C3)");
		sheet1.getCell("E6").setValue("=SUM(Sheet1!C3)");
		sheet1.getCell("F5").setValue("=SUM(Sheet1!C3,1)");
		
		Assert.assertEquals(3D, sheet1.getCell("C3").getValue());
		Assert.assertEquals(3D, sheet1.getCell("G1").getValue());
		Assert.assertEquals(3D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(3D, sheet1.getCell("E6").getValue());
		Assert.assertEquals(4D, sheet1.getCell("F5").getValue());
		
		sheet1.getCell("A1").setValue(5);
		Assert.assertEquals(5D, sheet1.getCell("C3").getValue());
		Assert.assertEquals(5D, sheet1.getCell("G1").getValue());
		Assert.assertEquals(5D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(5D, sheet1.getCell("E6").getValue());
		Assert.assertEquals(6D, sheet1.getCell("F5").getValue());
		
		sheet1.insertRow(1, 1);
		
		Assert.assertEquals("A1", sheet1.getCell("C4").getFormulaValue());
		Assert.assertEquals("C4", sheet1.getCell("G1").getFormulaValue());
		Assert.assertEquals("SUM(C4)", sheet1.getCell("A10").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!C4)", sheet1.getCell("E7").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!C4,1)", sheet1.getCell("F6").getFormulaValue());
		
		sheet1.getCell("A1").setValue(7);
		Assert.assertEquals(7D, sheet1.getCell("C4").getValue());
		Assert.assertEquals(7D, sheet1.getCell("G1").getValue());
		Assert.assertEquals(7D, sheet1.getCell("A10").getValue());
		Assert.assertEquals(7D, sheet1.getCell("E7").getValue());
		Assert.assertEquals(8D, sheet1.getCell("F6").getValue());
		
		sheet1.deleteRow(1, 2);
		Assert.assertEquals("A1", sheet1.getCell("C2").getFormulaValue());
		Assert.assertEquals("C2", sheet1.getCell("G1").getFormulaValue());
		Assert.assertEquals("SUM(C2)", sheet1.getCell("A8").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!C2)", sheet1.getCell("E5").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!C2,1)", sheet1.getCell("F4").getFormulaValue());
		
		sheet1.getCell("A1").setValue(9);
		Assert.assertEquals(9D, sheet1.getCell("C2").getValue());
		Assert.assertEquals(9D, sheet1.getCell("G1").getValue());
		Assert.assertEquals(9D, sheet1.getCell("A8").getValue());
		Assert.assertEquals(9D, sheet1.getCell("E5").getValue());
		Assert.assertEquals(10D, sheet1.getCell("F4").getValue());
		
		//trim
		sheet1.deleteRow(1, 1);
		Assert.assertEquals("#REF!", sheet1.getCell("G1").getFormulaValue());
		Assert.assertEquals("SUM(#REF!)", sheet1.getCell("A7").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!#REF!)", sheet1.getCell("E4").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!#REF!,1)", sheet1.getCell("F3").getFormulaValue());
	}
	
	@Test
	public void testFormulaShiftAfterInsertDeleteColumn(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("A1").setValue(3);
		sheet1.getCell("C3").setValue("=A1");
		sheet1.getCell("G1").setValue("=C3");
		sheet1.getCell("A9").setValue("=SUM(C3)");
		sheet1.getCell("E6").setValue("=SUM(Sheet1!C3)");
		sheet1.getCell("F5").setValue("=SUM(Sheet1!C3,1)");
		
		Assert.assertEquals(3D, sheet1.getCell("C3").getValue());
		Assert.assertEquals(3D, sheet1.getCell("G1").getValue());
		Assert.assertEquals(3D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(3D, sheet1.getCell("E6").getValue());
		Assert.assertEquals(4D, sheet1.getCell("F5").getValue());
		
		sheet1.getCell("A1").setValue(5);
		Assert.assertEquals(5D, sheet1.getCell("C3").getValue());
		Assert.assertEquals(5D, sheet1.getCell("G1").getValue());
		Assert.assertEquals(5D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(5D, sheet1.getCell("E6").getValue());
		Assert.assertEquals(6D, sheet1.getCell("F5").getValue());
		
		sheet1.insertColumn(1, 1);
		
		Assert.assertEquals("A1", sheet1.getCell("D3").getFormulaValue());
		Assert.assertEquals("D3", sheet1.getCell("H1").getFormulaValue());
		Assert.assertEquals("SUM(D3)", sheet1.getCell("A9").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!D3)", sheet1.getCell("F6").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!D3,1)", sheet1.getCell("G5").getFormulaValue());
		
		sheet1.getCell("A1").setValue(7);
		Assert.assertEquals(7D, sheet1.getCell("D3").getValue());
		Assert.assertEquals(7D, sheet1.getCell("H1").getValue());
		Assert.assertEquals(7D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(7D, sheet1.getCell("F6").getValue());
		Assert.assertEquals(8D, sheet1.getCell("G5").getValue());
		
		sheet1.deleteColumn(1, 2);
		Assert.assertEquals("A1", sheet1.getCell("B3").getFormulaValue());
		Assert.assertEquals("B3", sheet1.getCell("F1").getFormulaValue());
		Assert.assertEquals("SUM(B3)", sheet1.getCell("A9").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!B3)", sheet1.getCell("D6").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!B3,1)", sheet1.getCell("E5").getFormulaValue());
		
		sheet1.getCell("A1").setValue(9);
		Assert.assertEquals(9D, sheet1.getCell("B3").getValue());
		Assert.assertEquals(9D, sheet1.getCell("F1").getValue());
		Assert.assertEquals(9D, sheet1.getCell("A9").getValue());
		Assert.assertEquals(9D, sheet1.getCell("D6").getValue());
		Assert.assertEquals(10D, sheet1.getCell("E5").getValue());
		
		//trim
		sheet1.deleteColumn(1, 1);
		Assert.assertEquals("#REF!", sheet1.getCell("E1").getFormulaValue());
		Assert.assertEquals("SUM(#REF!)", sheet1.getCell("A9").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!#REF!)", sheet1.getCell("C6").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!#REF!,1)", sheet1.getCell("D5").getFormulaValue());
	}
	
	@Test
	public void testLinkingShift(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("C1").setValue(4);
		sheet1.getCell("E1").setValue("=C1");
		sheet1.getCell("D3").setValue("=SUM(E1)");
		Assert.assertEquals(4D, sheet1.getCell("E1").getValue());
		Assert.assertEquals(4D, sheet1.getCell("D3").getValue());
		
		sheet1.getCell("C3").setValue(3);
		sheet1.getCell("E3").setValue("=C3");
		sheet1.getCell("G3").setValue("=SUM(E3)");
		sheet1.getCell("G6").setValue("=SUM(Sheet1!G3,1)");
		sheet1.getCell("G1").setValue("=SUM(Sheet1!G3,1)");
		
		Assert.assertEquals(3D, sheet1.getCell("E3").getValue());
		Assert.assertEquals(3D, sheet1.getCell("G3").getValue());
		Assert.assertEquals(4D, sheet1.getCell("G6").getValue());
		Assert.assertEquals(4D, sheet1.getCell("G1").getValue());
		
		sheet1.deleteCell(new CellRegion("A3"), true);
		Assert.assertEquals("C1", sheet1.getCell("E1").getFormulaValue());
		Assert.assertEquals("SUM(E1)", sheet1.getCell("C3").getFormulaValue());
		sheet1.getCell("C1").setValue(6);
		Assert.assertEquals(6D, sheet1.getCell("E1").getValue());
		Assert.assertEquals(6D, sheet1.getCell("C3").getValue());
		
		Assert.assertEquals("B3", sheet1.getCell("D3").getFormulaValue());
		Assert.assertEquals("SUM(D3)", sheet1.getCell("F3").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!F3,1)", sheet1.getCell("G6").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!F3,1)", sheet1.getCell("G1").getFormulaValue());
		sheet1.getCell("B3").setValue(5);
		Assert.assertEquals(5D, sheet1.getCell("D3").getValue());
		Assert.assertEquals(5D, sheet1.getCell("F3").getValue());
		Assert.assertEquals(6D, sheet1.getCell("G6").getValue());
		Assert.assertEquals(6D, sheet1.getCell("G1").getValue());
		
		sheet1.deleteColumn(0, 0);
		Assert.assertEquals("B1", sheet1.getCell("D1").getFormulaValue());
		Assert.assertEquals("SUM(D1)", sheet1.getCell("B3").getFormulaValue());
		sheet1.getCell("B1").setValue(8);
		Assert.assertEquals(8D, sheet1.getCell("D1").getValue());
		Assert.assertEquals(8D, sheet1.getCell("B3").getValue());
		
		Assert.assertEquals("A3", sheet1.getCell("C3").getFormulaValue());
		Assert.assertEquals("SUM(C3)", sheet1.getCell("E3").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!E3,1)", sheet1.getCell("F6").getFormulaValue());
		Assert.assertEquals("SUM(Sheet1!E3,1)", sheet1.getCell("F1").getFormulaValue());
		sheet1.getCell("A3").setValue(7);
		Assert.assertEquals(7D, sheet1.getCell("C3").getValue());
		Assert.assertEquals(7D, sheet1.getCell("E3").getValue());
		Assert.assertEquals(8D, sheet1.getCell("F6").getValue());
		Assert.assertEquals(8D, sheet1.getCell("F1").getValue());
	}
	
	@Test
	public void testLinkedShift(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("A1").setValue(12);
		sheet1.getCell("A2").setValue("=A1");
		sheet1.getCell("A3").setValue("=A2");
		
		Assert.assertEquals(12D, sheet1.getCell("A1").getValue());
		Assert.assertEquals(12D, sheet1.getCell("A2").getValue());
		Assert.assertEquals(12D, sheet1.getCell("A3").getValue());
		
		sheet1.moveCell(new CellRegion("A1"),0,1);
		
		Assert.assertEquals("B1", sheet1.getCell("A2").getFormulaValue());
		Assert.assertEquals("A2", sheet1.getCell("A3").getFormulaValue());
		
		sheet1.getCell("B1").setValue(22);
		
		Assert.assertEquals(null, sheet1.getCell("A1").getValue());
		Assert.assertEquals(22D, sheet1.getCell("B1").getValue());
		Assert.assertEquals(22D, sheet1.getCell("A2").getValue());
		Assert.assertEquals(22D, sheet1.getCell("A3").getValue());
	}
	
	@Test
	public void testOverlapShift(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("B2").setValue(1);
		sheet1.getCell("B3").setValue(2);
		sheet1.getCell("C2").setValue(3);
		sheet1.getCell("C3").setValue(4);
		sheet1.getCell("C4").setValue(5);
		sheet1.getCell("D3").setValue(6);
		sheet1.getCell("D4").setValue(7);
		
		sheet1.getCell("F1").setValue("=SUM(A1:B2)");
		sheet1.getCell("F2").setValue("=SUM(B2:C3)");
		sheet1.getCell("F3").setValue("=SUM(C3:D4)");
		sheet1.getCell("G1").setValue("=A1");
		sheet1.getCell("G2").setValue("=B2");
		sheet1.getCell("G3").setValue("=C3");
		sheet1.getCell("G4").setValue("=D4");

		Assert.assertEquals(1D, sheet1.getCell("F1").getValue());
		Assert.assertEquals(10D, sheet1.getCell("F2").getValue());
		Assert.assertEquals(22D, sheet1.getCell("F3").getValue());
		Assert.assertEquals(0D, sheet1.getCell("G1").getValue());
		Assert.assertEquals(1D, sheet1.getCell("G2").getValue());
		Assert.assertEquals(4D, sheet1.getCell("G3").getValue());
		Assert.assertEquals(7D, sheet1.getCell("G4").getValue());
		
		sheet1.moveCell(new CellRegion("B2:C3"), -1,-1);
		
		Assert.assertEquals("SUM(#REF!)", sheet1.getCell("F1").getFormulaValue());
		Assert.assertEquals("SUM(A1:B2)", sheet1.getCell("F2").getFormulaValue());
		Assert.assertEquals("SUM(C3:D4)", sheet1.getCell("F3").getFormulaValue());
//		Assert.assertEquals("#REF!", sheet1.getCell("G1").getFormulaValue()); //TODO shouldn't fail here, or spec?
		Assert.assertEquals("A1", sheet1.getCell("G2").getFormulaValue());
		Assert.assertEquals("B2", sheet1.getCell("G3").getFormulaValue());
		Assert.assertEquals("D4", sheet1.getCell("G4").getFormulaValue());
		
		
		Assert.assertEquals("#REF!", sheet1.getCell("F1").getErrorValue().getErrorString());
		Assert.assertEquals(10D, sheet1.getCell("F2").getValue());
		Assert.assertEquals(18D, sheet1.getCell("F3").getValue());
//		Assert.assertEquals("#REF!", sheet1.getCell("G1").getErrorValue().getErrorString());//TODO shouldn't fail here, or spec?
		Assert.assertEquals(1D, sheet1.getCell("G2").getValue());
		Assert.assertEquals(4D, sheet1.getCell("G3").getValue());
		Assert.assertEquals(7D, sheet1.getCell("G4").getValue());
		
	}
	
	private SSheet prepareValidationSheet(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("B2").setValue("A");
		sheet1.getCell("B3").setValue("B");
		sheet1.getCell("B4").setValue("C");
		sheet1.getCell("B5").setValue("D");
		
		SDataValidation dv = sheet1.addDataValidation(new CellRegion("B8"));
		dv.setValidationType(ValidationType.LIST);
		dv.setFormula1("B2:B5");
		Assert.assertEquals(4,dv.getNumOfValue());
		Assert.assertEquals("A",dv.getValue(0));
		Assert.assertEquals("B",dv.getValue(1));
		Assert.assertEquals("C",dv.getValue(2));
		Assert.assertEquals("D",dv.getValue(3));
		Assert.assertEquals("B8",dv.getRegions().iterator().next().getReferenceString());
		return sheet1;
	}
//	@Test
	public void testValidationFormulaShift(){
		SSheet sheet1 = prepareValidationSheet();
		
		SDataValidation dv = sheet1.getDataValidation(0);
		
		//insert row
		sheet1.insertRow(2, 2);
		
		Assert.assertEquals("B9",dv.getRegions().iterator().next().getReferenceString());
		
		Assert.assertEquals(5,dv.getNumOfValue());
		Assert.assertEquals("A",dv.getValue(0));
		Assert.assertEquals(null,dv.getValue(1));
		Assert.assertEquals("B",dv.getValue(2));
		Assert.assertEquals("C",dv.getValue(3));
		Assert.assertEquals("D",dv.getValue(4));
		
		
		//delete row
		
		//insert column
		
		//delete column
		
		//insert cell
		
		//delete cell
		
		//move cell
		
		

	}
	
	
	@Test
	public void testRenameSheet(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		DependencyTable dt = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
		Set<Ref> refs;
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		SSheet sheet2 = initialDataGrid(book.createSheet("Sheet2"));
		
		sheet1.getCell("A1").setValue(12);
		sheet1.getCell("B1").setValue("=A1");
		
		sheet1.getCell("A2").setValue(34);
		sheet1.getCell("B2").setValue("=Sheet1!A2");
		
		sheet1.getCell("A3").setValue(56);
		sheet2.getCell("C3").setValue("=Sheet1!A3"); //sheet2 a1
		
		sheet2.getCell("A4").setValue(78);
		sheet2.getCell("C4").setValue("=A4");
		
		Assert.assertEquals(12D,sheet1.getCell("B1").getValue());
		Assert.assertEquals(34D,sheet1.getCell("B2").getValue());
		Assert.assertEquals(56D,sheet2.getCell("C3").getValue());
		Assert.assertEquals(78D,sheet2.getCell("C4").getValue());

		
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A1")));
		Assert.assertEquals(1, refs.size());
		Ref ref = refs.iterator().next();
		CellRegion refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("B1",refRegion.getReferenceString());
		
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A2")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("B2",refRegion.getReferenceString());
		
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A3")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("C3",refRegion.getReferenceString());
		
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A4")));
		Assert.assertEquals(0, refs.size());
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet2.getCell("A4")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("C4",refRegion.getReferenceString());
		
		book.setSheetName(sheet1, "SheetX");
		
		//shouldn't has old depedence on old sheet name
		refs = dt.getDependents(new RefImpl(book.getBookName(),"Sheet1",0,0));//a1
		Assert.assertEquals(0, refs.size());
		refs = dt.getDependents(new RefImpl(book.getBookName(),"Sheet1",1,0));//a2
		Assert.assertEquals(0, refs.size());

		refs = dt.getDependents(new RefImpl(book.getBookName(),"Sheet1",2,0));//a3
		Assert.assertEquals(0, refs.size());
		
		refs = dt.getDependents(new RefImpl(book.getBookName(),"Sheet1",3,0));//a4
		Assert.assertEquals(0, refs.size());
		
		//still has in sheet2
		refs = dt.getDependents(new RefImpl(book.getBookName(),"Sheet2",3,0));//a4
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("C4",refRegion.getReferenceString());
		
		
		//
		Assert.assertEquals("A1",sheet1.getCell("B1").getFormulaValue());
		Assert.assertEquals("SheetX!A2",sheet1.getCell("B2").getFormulaValue());
		Assert.assertEquals("SheetX!A3",sheet2.getCell("C3").getFormulaValue());
		Assert.assertEquals("A4",sheet2.getCell("C4").getFormulaValue());
		
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A1")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("B1",refRegion.getReferenceString());
		
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A2")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("B2",refRegion.getReferenceString());
		
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A3")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("C3",refRegion.getReferenceString());
		
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A4")));
		Assert.assertEquals(0, refs.size());
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet2.getCell("A4")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("C4",refRegion.getReferenceString());
		
		
		sheet1.getCell("A1").setValue(33);
		sheet1.getCell("A2").setValue(55);
		sheet1.getCell("A3").setValue(77);
		sheet2.getCell("A4").setValue(99);
		
		Assert.assertEquals(33D,sheet1.getCell("B1").getValue());
		Assert.assertEquals(55D,sheet1.getCell("B2").getValue());
		Assert.assertEquals(77D,sheet2.getCell("C3").getValue());
		Assert.assertEquals(99D,sheet2.getCell("C4").getValue());
		
		
	}
	
	
	@Test
	public void testDeleteSheet(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		DependencyTable dt = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
		Set<Ref> refs;
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		SSheet sheet2 = initialDataGrid(book.createSheet("Sheet2"));
		
		sheet1.getCell("A1").setValue(12);
		sheet1.getCell("B1").setValue("=A1");
		
		sheet1.getCell("A2").setValue(34);
		sheet1.getCell("B2").setValue("=Sheet1!A2");
		
		sheet1.getCell("A3").setValue(56);
		sheet2.getCell("C3").setValue("=Sheet1!A3"); //sheet2 a1
		
		sheet2.getCell("A4").setValue(78);
		sheet2.getCell("C4").setValue("=A4");
		
		Assert.assertEquals(12D,sheet1.getCell("B1").getValue());
		Assert.assertEquals(34D,sheet1.getCell("B2").getValue());
		Assert.assertEquals(56D,sheet2.getCell("C3").getValue());
		Assert.assertEquals(78D,sheet2.getCell("C4").getValue());

		
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A1")));
		Assert.assertEquals(1, refs.size());
		Ref ref = refs.iterator().next();
		CellRegion refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("B1",refRegion.getReferenceString());
		
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A2")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("B2",refRegion.getReferenceString());
		
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A3")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("C3",refRegion.getReferenceString());
		
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell("A4")));
		Assert.assertEquals(0, refs.size());
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet2.getCell("A4")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("C4",refRegion.getReferenceString());
		
		book.deleteSheet(sheet1);
		
		//shouldn't has old depedence on old sheet name
		refs = dt.getDependents(new RefImpl(book.getBookName(),"Sheet1",0,0));//a1
		Assert.assertEquals(0, refs.size());
		refs = dt.getDependents(new RefImpl(book.getBookName(),"Sheet1",1,0));//a2
		Assert.assertEquals(0, refs.size());

		refs = dt.getDependents(new RefImpl(book.getBookName(),"Sheet1",2,0));//a3
		Assert.assertEquals(0, refs.size());
		
		refs = dt.getDependents(new RefImpl(book.getBookName(),"Sheet1",3,0));//a4
		Assert.assertEquals(0, refs.size());
		
		//still has in sheet2
		refs = dt.getDependents(new RefImpl(book.getBookName(),"Sheet2",3,0));//sheet2 a4
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("C4",refRegion.getReferenceString());
		
		
		//poi parse can't suuport #REF!A3
		Assert.assertEquals("'#REF'!A3",sheet2.getCell("C3").getFormulaValue());
		Assert.assertEquals("A4",sheet2.getCell("C4").getFormulaValue());
		
		refs = dt.getDependents(new RefImpl((AbstractCellAdv)sheet2.getCell("A4")));
		Assert.assertEquals(1, refs.size());
		ref = refs.iterator().next();
		refRegion = new CellRegion(ref.getRow(),ref.getColumn(),ref.getLastRow(),ref.getLastColumn());
		Assert.assertEquals("C4",refRegion.getReferenceString());
		
		
		sheet2.getCell("A4").setValue(99);
		
		Assert.assertEquals("#REF!",sheet2.getCell("C3").getErrorValue().getErrorString());
		Assert.assertEquals(99D,sheet2.getCell("C4").getValue());
		
		
	}
}
