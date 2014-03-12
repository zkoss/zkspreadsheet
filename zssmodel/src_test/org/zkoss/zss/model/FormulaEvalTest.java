/* FormulaEvalTest.java

	Purpose:
		
	Description:
		
	History:
		Dec 9, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.model;

import java.io.Closeable;
import java.io.InputStream;
import java.text.MessageFormat;
import java.util.Locale;
import java.util.Set;

import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.util.Locales;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.ErrorValue;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SheetRegion;
import org.zkoss.zss.model.SCell.CellType;
import org.zkoss.zss.model.impl.AbstractBookSeriesAdv;
import org.zkoss.zss.model.impl.BookSeriesBuilderImpl;
import org.zkoss.zss.model.impl.NameRefImpl;
import org.zkoss.zss.model.impl.RefImpl;
import org.zkoss.zss.model.sys.EngineFactory;
import org.zkoss.zss.model.sys.dependency.DependencyTable;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.model.sys.formula.FormulaEngine;
import org.zkoss.zss.model.sys.formula.FormulaExpression;
import org.zkoss.zss.model.sys.formula.FormulaParseContext;
import org.zkoss.zss.range.SImporter;
import org.zkoss.zss.range.SRange;
import org.zkoss.zss.range.SRanges;
import org.zkoss.zss.range.impl.imexp.ExcelImportFactory;

/**
 * @author Pao
 */
public class FormulaEvalTest {

	private SImporter importer;
	@BeforeClass
	static public void beforeClass() {
		Setup.touch();
	}
	@Before
	public void beforeTest() {
		importer= new ExcelImportFactory().createImporter();
		Locales.setThreadLocal(Locale.TAIWAN);
	}

	@Test
	public void testBasicEvaluation() {
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = book.createSheet("Sheet1");
		sheet1.getCell(0, 0).setFormulaValue("SUM(C1:C10)");
		sheet1.getCell(0, 1).setNumberValue(55.0);
		sheet1.getCell(1, 0).setFormulaValue("AVERAGE(Sheet2!C1:C10)");
		sheet1.getCell(1, 1).setNumberValue(5.5);
		for(int r = 0; r < 10; ++r) {
			sheet1.getCell(r, 2).setValue(r + 1);
		}
		SSheet sheet2 = book.createSheet("Sheet2");
		for(int r = 0; r < 10; ++r) {
			sheet2.getCell(r, 2).setValue(r + 1);
		}

		// evaluation
		for(int r = 0; r < 2; ++r) {
			Assert.assertEquals("row " + r, sheet1.getCell(r, 1).getValue().toString(), sheet1.getCell(r, 0)
					.getValue().toString());
		}
	}

	@Test
	public void testBasicEvaluation2() {
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = book.createSheet("Sheet1");
		SSheet sheet2 = book.createSheet("Sheet2");
		SSheet sheetA = book.createSheet("SheetA");

		// cells in sheet1
		// basic formula
		sheet1.getCell(0, 0).setFormulaValue("B1");
		sheet1.getCell(1, 0).setFormulaValue(" B2");
		sheet1.getCell(2, 0).setFormulaValue("SUM(C1:C10)");
		sheet1.getCell(3, 0).setFormulaValue("A3 / 2");
		sheet1.getCell(4, 0).setFormulaValue("A3 + A4");
		sheet1.getCell(5, 0).setFormulaValue("AVERAGE(C:C)");
		sheet1.getCell(6, 0).setFormulaValue("D1");
		// sheet ref.
		sheet1.getCell(7, 0).setFormulaValue("Sheet2!D2");
		sheet1.getCell(8, 0).setFormulaValue("AVERAGE(Sheet2!D1:D10)");
		// custom function
		// sheet1.getCell(9, 0).setFormulaValue(" TEST() "); // FIXME
		// name range
		sheet1.getCell(9, 0).setFormulaValue("SUM(ABC)");
		sheet1.getCell(10, 0).setFormulaValue("SUM(DEF)");
		// 3D ref.
		sheet1.getCell(11, 0).setFormulaValue("SUM(Sheet1:SheetA!D1:E10)");
		sheet1.getCell(12, 0).setFormulaValue("SUM(Sheet1:SheetA!D2)");

		// error test
		sheet1.getCell(13, 0).setFormulaValue("SUM(NotExistedName)");

		sheet1.getCell(0, 1).setStringValue("HELLO!!!");
		sheet1.getCell(1, 1).setStringValue("ZK!!!");
		for(int i = 0; i < 10; ++i) {
			sheet1.getCell(i, 2).setNumberValue((double)i); // B
			sheet1.getCell(i, 3).setNumberValue(i * 10.0); // C
			sheet1.getCell(i, 4).setNumberValue(i * 10.0); // D
		}

		// sheet2
		// cells in sheet2
		for(int i = 0; i < 10; ++i) {
			sheet2.getCell(i, 3).setNumberValue(i * 10.0); // C
			sheet2.getCell(i, 4).setNumberValue(i * 10.0); // D
		}

		// sheetA
		// cells in sheetAd
		for(int i = 0; i < 10; ++i) {
			sheetA.getCell(i, 3).setNumberValue(i * 10.0); // C
			sheetA.getCell(i, 4).setNumberValue(i * 10.0); // D
		}

		// named range
		// book.createName("ABC", "Sheet1").setRefersToFormula("C1:C10"); // TODO ZSS 3.5
		book.createName("ABC").setRefersToFormula("Sheet1!C1:C10");
		book.createName("DEF").setRefersToFormula("SheetA!D1:D10");
		book.createName("GHI").setRefersToFormula("Sheet1:Sheet2!C1:C10");

		// test
		String[] expected = new String[]{"HELLO!!!", "ZK!!!", "45.0", "22.5", "67.5", "4.5", "0.0", "10.0",
				"45.0", "45.0", "450.0", (450 * 6) + ".0", "30.0", "#NAME?"};
		for(int i = 0; i <= 99; ++i) {
			SCell cell = sheet1.getCell(i, 0);
			if(cell.isNull()) {
				break;
			}

			switch(cell.getFormulaResultType()) {
				case STRING:
					Assert.assertEquals(expected[i], cell.getStringValue());
					break;
				case NUMBER:
					Assert.assertEquals(expected[i], cell.getNumberValue().toString());
					break;
				case ERROR:
					Assert.assertEquals(expected[i], cell.getErrorValue().getErrorString());
					break;
				default:
					Assert.fail("not expected result type: " + cell.getFormulaResultType());
					break;
			}
		}
	}

	@Test
	public void testFormulaDependencyTracking() {

		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = book.createSheet("Sheet1");
		book.createSheet("Sheet2");
		book.createSheet("Sheet3");
		DependencyTable table = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();

		Assert.assertTrue(table.getDependents(toAreaRef("Sheet1", null, "A2:C2")).isEmpty());
		book.createName("ABC").setRefersToFormula("Sheet1!A1:A5");
		book.createName("DEF").setRefersToFormula("Sheet1:Sheet2!B2");
		book.createName("GHI").setRefersToFormula("Sheet1:Sheet3!C2:C3");

		Set<Ref> dependents = table.getDependents(toAreaRef("Sheet1", null, "A2:C2"));
		Assert.assertEquals(dependents.toString(), 3, dependents.size());
		Assert.assertTrue(dependents.contains(toNameRef("ABC")));
		Assert.assertTrue(dependents.contains(toNameRef("DEF")));
		Assert.assertTrue(dependents.contains(toNameRef("GHI")));

		sheet1.getCell(0, 0).setFormulaValue("B1");
		sheet1.getCell(1, 0).setFormulaValue("SUM(C1:C10)");
		sheet1.getCell(2, 0).setFormulaValue("B3 / 2");
		sheet1.getCell(3, 0).setFormulaValue("B3 + B4");
		sheet1.getCell(4, 0).setFormulaValue("AVERAGE(C:C)");
		sheet1.getCell(5, 0).setFormulaValue("D1");
		sheet1.getCell(6, 0).setFormulaValue("Sheet2!D2");
		sheet1.getCell(7, 0).setFormulaValue("AVERAGE(Sheet2!D1:D10)");
		sheet1.getCell(8, 0).setFormulaValue("SUM(ABC)");
		sheet1.getCell(9, 0).setFormulaValue("SUM(NotExistedName)");
		sheet1.getCell(10, 0).setFormulaValue("SUM(Sheet1:Sheet3!D1:E10)");
		sheet1.getCell(11, 0).setFormulaValue("SUM(Sheet1:Sheet3!D2)");

		dependents = table.getDependents(toAreaRef("Sheet1", null, "A5"));
		Assert.assertEquals(dependents.toString(), 2, dependents.size());
		Assert.assertTrue(dependents.toString(), dependents.contains(toNameRef("ABC")));
		Assert.assertTrue(dependents.toString(), dependents.contains(toCellRef("Sheet1", null, "A9")));

		dependents = table.getDependents(toAreaRef("Sheet1", null, "B4"));
		Assert.assertEquals(dependents.toString(), 3, dependents.size());
		Assert.assertTrue(dependents.toString(), dependents.contains(toCellRef("Sheet1", null, "A4")));
		Assert.assertTrue(dependents.toString(), dependents.contains(toNameRef("ABC")));
		Assert.assertTrue(dependents.toString(), dependents.contains(toCellRef("Sheet1", null, "A9")));

		dependents = table.getDependents(toAreaRef("Sheet2", null, "D2"));
		Assert.assertEquals(dependents.toString(), 4, dependents.size());
		Assert.assertTrue(dependents.toString(), dependents.contains(toCellRef("Sheet1", null, "A7")));
		Assert.assertTrue(dependents.toString(), dependents.contains(toCellRef("Sheet1", null, "A8")));
		Assert.assertTrue(dependents.toString(), dependents.contains(toCellRef("Sheet1", null, "A11")));
		Assert.assertTrue(dependents.toString(), dependents.contains(toCellRef("Sheet1", null, "A12")));
	}
	
	@Test
	public void testFormulaDependencyTrackingWithExternalBook() {

		SBook book1 = SBooks.createBook("book1");
		SSheet sheetA = book1.createSheet("SheetA");
		SBook book2 = SBooks.createBook("book2");
		SSheet sheetB = book2.createSheet("SheetB");
		getCell(sheetA, "A1").setFormulaValue("A2 + 1");
		getCell(sheetA, "A2").setNumberValue(1.0);
		getCell(sheetA, "A3").setFormulaValue("SUM(NameA)");
		getCell(sheetA, "A4").setFormulaValue("SUM(NameB)");
		getCell(sheetB, "B1").setFormulaValue("B2 + 1");
		getCell(sheetB, "B2").setNumberValue(2.0);
		getCell(sheetB, "B3").setFormulaValue("SUM(NameB)");
		getCell(sheetB, "B4").setFormulaValue("SUM(NameA)");
		book1.createName("NameA").setRefersToFormula("SheetA!A1:A2");
		book2.createName("NameB").setRefersToFormula("SheetB!B1:B2");

		// test initial
		Assert.assertEquals(2.0, getCell(sheetA, "A1").getNumberValue(), 0.0001);
		Assert.assertEquals(1.0, getCell(sheetA, "A2").getNumberValue(), 0.0001);
		Assert.assertEquals(3.0, getCell(sheetA, "A3").getNumberValue(), 0.0001);
		Assert.assertEquals(3.0, getCell(sheetB, "B1").getNumberValue(), 0.0001);
		Assert.assertEquals(2.0, getCell(sheetB, "B2").getNumberValue(), 0.0001);
		Assert.assertEquals(5.0, getCell(sheetB, "B3").getNumberValue(), 0.0001);
		Assert.assertEquals(ErrorValue.INVALID_NAME, getCell(sheetA, "A4").getErrorValue().getCode());
		Assert.assertEquals(ErrorValue.INVALID_NAME, getCell(sheetB, "B4").getErrorValue().getCode());
		
		DependencyTable table1 = ((AbstractBookSeriesAdv)book1.getBookSeries()).getDependencyTable();
		Set<Ref> dependents = table1.getDependents(toCellRef("book1", "SheetA", null, "A2"));
		Assert.assertEquals(dependents.toString(), 3, dependents.size());
		Assert.assertTrue(dependents.contains(new NameRefImpl("book1", null, "NameA")));
		Assert.assertTrue(dependents.contains(toCellRef("book1", "SheetA", null, "A1")));
		Assert.assertTrue(dependents.contains(toCellRef("book1", "SheetA", null, "A3")));
		
		DependencyTable table2 = ((AbstractBookSeriesAdv)book2.getBookSeries()).getDependencyTable();
		dependents = table2.getDependents(toCellRef("book2", "SheetB", null, "B2"));
		Assert.assertEquals(dependents.toString(), 3, dependents.size());
		Assert.assertTrue(dependents.contains(new NameRefImpl("book2", null, "NameB")));
		Assert.assertTrue(dependents.contains(toCellRef("book2", "SheetB", null, "B1")));
		Assert.assertTrue(dependents.contains(toCellRef("book2", "SheetB", null, "B3")));

		// test dependency merge
		// we can get other book's dependency after built book series
		dependents = table1.getDependents(toCellRef("book2", "SheetB", null, "B2"));
		Assert.assertEquals(dependents.toString(), 0, dependents.size());
		Assert.assertFalse(dependents.contains(new NameRefImpl("book2", null, "NameB")));
		Assert.assertFalse(dependents.contains(toCellRef("book2", "SheetB", null, "B1")));
		Assert.assertFalse(dependents.contains(toCellRef("book2", "SheetB", null, "B3")));
		dependents = table2.getDependents(toCellRef("book1", "SheetA", null, "A2"));
		Assert.assertEquals(dependents.toString(), 0, dependents.size());
		Assert.assertFalse(dependents.contains(new NameRefImpl("book1", null, "NameA")));
		Assert.assertFalse(dependents.contains(toCellRef("book1", "SheetA", null, "A1")));
		Assert.assertFalse(dependents.contains(toCellRef("book1", "SheetA", null, "A3")));

		new BookSeriesBuilderImpl().buildBookSeries(book1, book2);
		table1 = ((AbstractBookSeriesAdv)book1.getBookSeries()).getDependencyTable();
		table2 = ((AbstractBookSeriesAdv)book2.getBookSeries()).getDependencyTable();
		dependents = table1.getDependents(toCellRef("book2", "SheetB", null, "B2"));
		Assert.assertEquals(dependents.toString(), 3, dependents.size());
		Assert.assertTrue(dependents.contains(new NameRefImpl("book2", null, "NameB")));
		Assert.assertTrue(dependents.contains(toCellRef("book2", "SheetB", null, "B1")));
		Assert.assertTrue(dependents.contains(toCellRef("book2", "SheetB", null, "B3")));
		dependents = table2.getDependents(toCellRef("book1", "SheetA", null, "A2"));
		Assert.assertEquals(dependents.toString(), 3, dependents.size());
		Assert.assertTrue(dependents.contains(new NameRefImpl("book1", null, "NameA")));
		Assert.assertTrue(dependents.contains(toCellRef("book1", "SheetA", null, "A1")));
		Assert.assertTrue(dependents.contains(toCellRef("book1", "SheetA", null, "A3")));
		
		// test dependency from external books reference
		getCell(sheetA, "A5").setFormulaValue("[book2]SheetB!B5 + 1");
		getCell(sheetB, "B5").setNumberValue(3.0);
		Assert.assertEquals(4.0, getCell(sheetA, "A5").getNumberValue(), 0.0001);
		Assert.assertEquals(3.0, getCell(sheetB, "B5").getNumberValue(), 0.0001);

		dependents = table1.getDependents(toCellRef("book2", "SheetB", null, "B5"));
		Assert.assertEquals(dependents.toString(), 1, dependents.size());
		Assert.assertTrue(dependents.contains(toCellRef("book1", "SheetA", null, "A5")));
		dependents = table2.getDependents(toCellRef("book2", "SheetB", null, "B5"));
		Assert.assertEquals(dependents.toString(), 1, dependents.size());
		Assert.assertTrue(dependents.contains(toCellRef("book1", "SheetA", null, "A5")));

		getCell(sheetA, "A6").setNumberValue(4.0);
		getCell(sheetB, "B6").setFormulaValue("[book1]SheetA!A6 + 1");
		Assert.assertEquals(4.0, getCell(sheetA, "A6").getNumberValue(), 0.0001);
		Assert.assertEquals(5.0, getCell(sheetB, "B6").getNumberValue(), 0.0001);

		dependents = table1.getDependents(toCellRef("book1", "SheetA", null, "A6"));
		Assert.assertEquals(dependents.toString(), 1, dependents.size());
		Assert.assertTrue(dependents.contains(toCellRef("book2", "SheetB", null, "B6")));
		dependents = table2.getDependents(toCellRef("book1", "SheetA", null, "A6"));
		Assert.assertEquals(dependents.toString(), 1, dependents.size());
		Assert.assertTrue(dependents.contains(toCellRef("book2", "SheetB", null, "B6")));

		// extra test
		Assert.assertEquals(ErrorValue.INVALID_NAME, getCell(sheetA, "A4").getErrorValue().getCode());
		Assert.assertEquals(ErrorValue.INVALID_NAME, getCell(sheetB, "B4").getErrorValue().getCode());
	}
	
	private SCell getCell(SSheet sheet, String ref) {
		CellRegion region = new CellRegion(ref);
		return sheet.getCell(region.getRow(), region.getColumn());
	}

	private Ref toAreaRef(String sheet1, String sheet2, String area) {
		return toAreaRef("book1", sheet1, sheet2, area);
	}
	private Ref toAreaRef(String book, String sheet1, String sheet2, String area) {
		CellRegion r = new CellRegion(area);
		return new RefImpl(book, sheet1, sheet2, r.getRow(), r.getColumn(), r.getLastRow(),
				r.getLastColumn());
	}

	private Ref toSheetRef(String sheet1) {
		return new RefImpl("book1", sheet1);
	}

	private Ref toCellRef(String sheet1, String sheet2, String cell) {
		return toCellRef("book1", sheet1, sheet2, cell);
	}
	private Ref toCellRef(String book, String sheet1, String sheet2, String cell) {
		CellRegion r = new CellRegion(cell);
		return new RefImpl(book, sheet1, sheet2, r.getRow(), r.getColumn());
	}

	private Ref toNameRef(String name) {
		return new NameRefImpl("book1", null, name);
	}
	
	
	@Test
	public void testEvalAndModifyNormal(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		DependencyTable table = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
		
		SSheet sheet1 = book.createSheet("Sheet1");
		SSheet sheet2 = book.createSheet("Sheet2");

		SCell cell = sheet1.getCell(0, 0);
		cell.setFormulaValue("Sheet2!A1");
		
		Assert.assertEquals(CellType.FORMULA, cell.getType());
		Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
		Assert.assertEquals(0D, cell.getValue());
		
		sheet2.getCell(0, 0).setValue("ABC");
		Assert.assertEquals(CellType.STRING, cell.getFormulaResultType());
		Assert.assertEquals("ABC", cell.getValue());
		
		sheet2.getCell(0, 0).setValue(Boolean.TRUE);
		Assert.assertEquals(CellType.BOOLEAN, cell.getFormulaResultType());
		Assert.assertEquals(true, cell.getValue());
		
		sheet2.getCell(0, 0).setValue(123);
		Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
		Assert.assertEquals(123D, cell.getValue());
		
		Set<Ref> dependents = table.getDependents(toCellRef("Sheet2",null,"A1"));
		Assert.assertEquals(1, dependents.size());
		Ref ref = dependents.iterator().next();
		Assert.assertEquals("Sheet1", ref.getSheetName());
		Assert.assertEquals(null, ref.getLastSheetName());
		Assert.assertEquals(0, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(0, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());
		
		
		//use range, it notify to clean automatically
		cell = sheet1.getCell(1, 0);
		cell.setFormulaValue("Sheet2!A2");
		SRange r = SRanges.range(sheet2, 1, 0);
		
		Assert.assertEquals(CellType.FORMULA, cell.getType());
		Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
		Assert.assertEquals(0D, cell.getValue());
		
		r.setValue("ABC");
		Assert.assertEquals(CellType.STRING, cell.getFormulaResultType());
		Assert.assertEquals("ABC", cell.getValue());
		
		r.setValue(Boolean.TRUE);
		Assert.assertEquals(CellType.BOOLEAN, cell.getFormulaResultType());
		Assert.assertEquals(true, cell.getValue());
		
		r.setValue(123);
		Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
		Assert.assertEquals(123D, cell.getValue());
		
		dependents = table.getDependents(toCellRef("Sheet2",null,"A2"));
		Assert.assertEquals(1, dependents.size());
		ref = dependents.iterator().next();
		Assert.assertEquals("Sheet1", ref.getSheetName());
		Assert.assertEquals(null, ref.getLastSheetName());
		Assert.assertEquals(1, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(1, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());
		
		//TODO test clear
		cell.setFormulaValue("Sheet2!A3");
		
		dependents = table.getDependents(toCellRef("Sheet2",null,"A3"));
		Assert.assertEquals(1, dependents.size());
		ref = dependents.iterator().next();
		Assert.assertEquals("Sheet1", ref.getSheetName());
		Assert.assertEquals(null, ref.getLastSheetName());
		Assert.assertEquals(1, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(1, ref.getLastRow());
		
		//should get 0
		dependents = table.getDependents(toCellRef("Sheet2",null,"A2"));
		Assert.assertEquals(0, dependents.size());
	}
	
	@Test
	public void testEvalNoRef(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		DependencyTable table = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
		
		SSheet sheet1 = book.createSheet("Sheet1");

		SCell cell = sheet1.getCell(0, 0);
		cell.setFormulaValue("Sheet2!A1");
		
		Assert.assertEquals(CellType.FORMULA, cell.getType());
		Assert.assertEquals(CellType.ERROR, cell.getFormulaResultType());
		Assert.assertEquals("#REF!", cell.getErrorValue().getErrorString());
		
		
		Set<Ref> dependents = table.getDependents(toSheetRef("Sheet2"));
		Assert.assertEquals(1, dependents.size());
		Ref ref = dependents.iterator().next();
		Assert.assertEquals("Sheet1", ref.getSheetName());
		Assert.assertEquals(null, ref.getLastSheetName());
		Assert.assertEquals(0, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(0, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());
		
		
		SSheet sheet2 = book.createSheet("Sheet2");
		//should clear all for any sheet state change.
		
		Assert.assertEquals(CellType.FORMULA, cell.getType());
		Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
		Assert.assertEquals(0D, cell.getValue());
		
		sheet2.getCell(0, 0).setValue("ABC");
		Assert.assertEquals(CellType.STRING, cell.getFormulaResultType());
		Assert.assertEquals("ABC", cell.getValue());
	}

	@Test
	public void testClearFormulaDependency() {
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = book.createSheet("Sheet1");
		DependencyTable table = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();

		// initial test
		Assert.assertTrue(table.getDependents(toAreaRef("Sheet1", null, "B1")).isEmpty());
		Assert.assertTrue(table.getDependents(toAreaRef("Sheet1", null, "C1")).isEmpty());
		// group2
		Assert.assertTrue(table.getDependents(toAreaRef("Sheet1", null, "B2")).isEmpty());
		Assert.assertTrue(table.getDependents(toAreaRef("Sheet1", null, "C2")).isEmpty());

		// make some dependencies
		sheet1.getCell(0, 0).setFormulaValue("B1 + C1"); // A1 = B1 + C1
		sheet1.getCell(0, 1).setNumberValue(1.0); // B1 = 1.0
		sheet1.getCell(0, 2).setNumberValue(2.0); // C1 = 2.0
		// group2, clean must not effect these
		sheet1.getCell(1, 0).setFormulaValue("B2 + C2"); // A2 = B2 + C2
		sheet1.getCell(1, 1).setNumberValue(3.0); // B2 = 3.0
		sheet1.getCell(1, 2).setNumberValue(4.0); // C2 = 4.0

		// test value
		Assert.assertEquals(Double.valueOf(3.0), sheet1.getCell(0, 0).getNumberValue());
		Assert.assertEquals(Double.valueOf(7.0), sheet1.getCell(1, 0).getNumberValue());

		// test normal dependencies
		Set<Ref> dependents;
		dependents = table.getDependents(toAreaRef("Sheet1", null, "B1"));
		Assert.assertEquals(dependents.toString(), 1, dependents.size());
		Assert.assertTrue(dependents.contains(toCellRef("Sheet1", null, "A1")));
		dependents = table.getDependents(toAreaRef("Sheet1", null, "C1"));
		Assert.assertEquals(dependents.toString(), 1, dependents.size());
		Assert.assertTrue(dependents.contains(toCellRef("Sheet1", null, "A1")));
		// group2
		dependents = table.getDependents(toAreaRef("Sheet1", null, "B2"));
		Assert.assertEquals(dependents.toString(), 1, dependents.size());
		Assert.assertTrue(dependents.contains(toCellRef("Sheet1", null, "A2")));
		dependents = table.getDependents(toAreaRef("Sheet1", null, "C2"));
		Assert.assertEquals(dependents.toString(), 1, dependents.size());
		Assert.assertTrue(dependents.contains(toCellRef("Sheet1", null, "A2")));

		// modify cell and make clear dependencies
		sheet1.getCell(0, 0).clearValue();
		Assert.assertEquals(CellType.BLANK, sheet1.getCell(0, 0).getType());
		Assert.assertEquals(Double.valueOf(7.0), sheet1.getCell(1, 0).getNumberValue());

		// test dependency clearing
		Assert.assertTrue(table.getDependents(toAreaRef("Sheet1", null, "B1")).isEmpty());
		Assert.assertTrue(table.getDependents(toAreaRef("Sheet1", null, "C1")).isEmpty());
		// group2
		dependents = table.getDependents(toAreaRef("Sheet1", null, "B2"));
		Assert.assertEquals(dependents.toString(), 1, dependents.size());
		Assert.assertTrue(dependents.contains(toCellRef("Sheet1", null, "A2")));
		dependents = table.getDependents(toAreaRef("Sheet1", null, "C2"));
		Assert.assertEquals(dependents.toString(), 1, dependents.size());
		Assert.assertTrue(dependents.contains(toCellRef("Sheet1", null, "A2")));

	}

	public SBook getBook(String path, String bookName) {
		InputStream is = null;
		try {
			is = FormulaEvalTest.class.getResourceAsStream(path);
			return importer.imports(is, bookName);
		} catch(Exception e) {
			e.printStackTrace();
		} finally {
			close(is);
		}
		return null;
	}

	public void close(Closeable r) {
		try {
			r.close();
		} catch(Exception e) {
		}
	}

	@Test
	public void testArrayValue() {
		SBook book = getBook("book/formula-eval.xlsx", "Book1");
		Assert.assertNotNull(book);
		SSheet sheet = book.getSheetByName("array");
		Assert.assertNotNull(sheet);

		// check formula
		int c = 2;
		for(int r = 1; r <= 6; ++r) { // C2:C7
			SCell cell = sheet.getCell(r, c);
			Assert.assertEquals("A2:A6", cell.getFormulaValue());
		}

		// check value
		c = 0;
		for(int r = 1; r <= 5; ++r) { // A2:A6
			SCell cell = sheet.getCell(r, c);
			Assert.assertEquals(r, cell.getNumberValue().intValue());
		}
		c = 2;
		for(int r = 1; r <= 5; ++r) { // C2:C6
			SCell cell = sheet.getCell(r, c);
			Assert.assertEquals(r, cell.getNumberValue().intValue());
		}

		// check C7
		SCell C7 = sheet.getCell(6, 2);
		Assert.assertEquals(CellType.ERROR, C7.getFormulaResultType());
		Assert.assertEquals(ErrorValue.INVALID_VALUE, C7.getErrorValue().getCode());
	}
	
	@Test
	@Ignore
	public void testArrayFormula() {
		SBook book = getBook("book/formula-eval.xlsx", "Book1");
		SSheet sheet = book.getSheetByName("array");
		Assert.assertNotNull(book);
		Assert.assertNotNull(sheet);

		// check eval. value
		SCell e2 = sheet.getCell(1, 4);
		Assert.assertFalse(e2.isNull());
		Assert.assertEquals(CellType.NUMBER, e2.getFormulaResultType());
		Assert.assertEquals(11, e2.getNumberValue().intValue());
	}
	
	@Test
	public void testExternalBookReference() {
		SBook book1, book2, book3;
		SSheet sheet1, sheet2;
		
		// direct creation
		book1 = SBooks.createBook("Book1");
		book2 = SBooks.createBook("Book2");
		book1.createSheet("external").getCell(0, 1).setFormulaValue("SUM(3 + [Book2]ref!C6 - 3)"); // B1
		book2.createSheet("ref").getCell(5, 2).setNumberValue(5.0);	// C6
		testExternalBookReference(book1, book2);
		
		// direct creation - complex 
		book1 = SBooks.createBook("Book1");
		book2 = SBooks.createBook("Book2");
		book3 = SBooks.createBook("Book3");
		sheet1 = book1.createSheet("external");
		sheet1.getCell(0, 1).setFormulaValue(" SUM(external:another!K1) + SUM( [Book2]ref!A1:A2 , [Book3]sheet1:sheet3!C1 ) + B2 - another!B3 "); // [Book1]external!B1
		// same book + same sheet
		sheet1.getCell(1, 1).setNumberValue(1.0); // [Book1]external!B2
		// same book + other sheet
		sheet2 = book1.createSheet("another");
		sheet2.getCell(2, 1).setNumberValue(1.0); // [Book1]another!B3
		// external book + area ref.
		sheet1 = book2.createSheet("ref");
		sheet1.getCell(0, 0).setNumberValue(1.0); // [Book2]ref!A1
		sheet1.getCell(1, 0).setNumberValue(1.0); // [Book2]ref!A2
		// external book + 3D ref.
		book3.createSheet("sheet1").getCell(0, 2).setNumberValue(1.0); // [Book3]sheet1!C1 
		book3.createSheet("sheet2").getCell(0, 2).setNumberValue(1.0); // [Book3]sheet2!C1
		book3.createSheet("sheet3").getCell(0, 2).setNumberValue(1.0); // [Book3]sheet3!C1
		testExternalBookReference(book1, book2, book3);

		// 2007
		book1 = getBook("book/formula-eval-external.xlsx", "Book1");
		book2 = getBook("book/formula-eval.xlsx", "formula-eval.xlsx");
		testExternalBookReference(book1, book2);

		// 2003
		book1 = getBook("book/formula-eval-external.xls", "Book1");
		book2 = getBook("book/formula-eval.xls", "formula-eval.xls");
		testExternalBookReference(book1, book2);
	}

	public void testExternalBookReference(SBook... books) {
		Assert.assertTrue(books.length >= 2);
		SBook book1 = books[0];

		// update book series
		for(SBook book : books) {
			Assert.assertNotNull(book);
		}
		new BookSeriesBuilderImpl().buildBookSeries(books); 

		// check evaluated value of B1
		SSheet sheet = book1.getSheetByName("external");
		SCell b1 = sheet.getCell(0, 1);
		Assert.assertFalse(b1.isNull());
		Assert.assertEquals(CellType.NUMBER, b1.getFormulaResultType());
		Assert.assertEquals(5, b1.getNumberValue().intValue());
	}
	
	@Test
	public void testFormulaMove() {

		FormulaEngine engine = EngineFactory.getInstance().createFormulaEngine();
		SBook book1 = SBooks.createBook("Book1");
		SSheet sheetA = book1.createSheet("SheetA");
		SSheet sheetB = book1.createSheet("SheetB");
		SBook book2 = SBooks.createBook("Book2");
		SSheet book2SheetA = book2.createSheet("SheetA");
		book2.createSheet("SheetB");
		new BookSeriesBuilderImpl().buildBookSeries(book1, book2);
		
		final int maxRow = book1.getMaxRowSize();
		final int maxColumn = book1.getMaxColumnSize();

		// shift rows
		String f = "SUM(C3:E5)+SUM(SheetA!C3:E5)+SUM(SheetB!C3:E5)";
		testFormulaMove(engine, sheetA, f, new SheetRegion(sheetA, "C3:E5"), 2, 0,
				"SUM(C5:E7)+SUM(SheetA!C5:E7)+SUM(SheetB!C3:E5)");
		testFormulaMove(engine, sheetA, f, new SheetRegion(sheetA, "C3:E5"), -2, 0,
				"SUM(C1:E3)+SUM(SheetA!C1:E3)+SUM(SheetB!C3:E5)");
		// shift columns
		testFormulaMove(engine, sheetA, f, new SheetRegion(sheetA, "C3:E5"), 0, 2,
				"SUM(E3:G5)+SUM(SheetA!E3:G5)+SUM(SheetB!C3:E5)");
		testFormulaMove(engine, sheetA, f, new SheetRegion(sheetA, "C3:E5"), 0, -2,
				"SUM(A3:C5)+SUM(SheetA!A3:C5)+SUM(SheetB!C3:E5)");
		// shift both
		testFormulaMove(engine, sheetA, f, new SheetRegion(sheetA, "C3:E5"), 2, 2,
				"SUM(E5:G7)+SUM(SheetA!E5:G7)+SUM(SheetB!C3:E5)");
		testFormulaMove(engine, sheetA, f, new SheetRegion(sheetA, "C3:E5"), -2, -2,
				"SUM(A1:C3)+SUM(SheetA!A1:C3)+SUM(SheetB!C3:E5)");
		
		// shift other sheet's region
		testFormulaMove(engine, sheetA, f, new SheetRegion(sheetB, "C3:E5"), 2, 2,
				"SUM(C3:E5)+SUM(SheetA!C3:E5)+SUM(SheetB!E5:G7)");
		testFormulaMove(engine, sheetA, f, new SheetRegion(sheetB, "C3:E5"), -2, -2,
				"SUM(C3:E5)+SUM(SheetA!C3:E5)+SUM(SheetB!A1:C3)");
		
		// out of bound, exceed min.
		testFormulaMove(engine, sheetA, f, new SheetRegion(sheetA, "C3:E5"), -3, 0,
				"SUM(C1:E2)+SUM(SheetA!C1:E2)+SUM(SheetB!C3:E5)");
		testFormulaMove(engine, sheetA, f, new SheetRegion(sheetA, "C3:E5"), 0, -3,
				"SUM(A3:B5)+SUM(SheetA!A3:B5)+SUM(SheetB!C3:E5)");
		// out of bound, exceed max.
		testFormulaMove(engine, sheetA, f, new SheetRegion(sheetA, "C3:E5"), maxRow-3, 0,
				MessageFormat.format("SUM(C{0}:E{0})+SUM(SheetA!C{0}:E{0})+SUM(SheetB!C3:E5)", String.valueOf(maxRow))); // only last row exceeds 
		testFormulaMove(engine, sheetA, f, new SheetRegion(sheetA, "C3:E5"), maxRow, 0,
				"SUM(#REF!)+SUM(SheetA!#REF!)+SUM(SheetB!C3:E5)"); // first and last both exceed 
		testFormulaMove(engine, sheetA, f, new SheetRegion(sheetA, "C3:E5"), 0, maxColumn-3,
				MessageFormat.format("SUM({0}:{1})+SUM(SheetA!{0}:{1})+SUM(SheetB!C3:E5)", new CellRegion(2, maxColumn-1).getReferenceString(), new CellRegion(4, maxColumn-1).getReferenceString())); // only last column exceeds 
		testFormulaMove(engine, sheetA, f, new SheetRegion(sheetA, "C3:E5"), 0, maxColumn,
				"SUM(#REF!)+SUM(SheetA!#REF!)+SUM(SheetB!C3:E5)"); // first and last both exceed 
		
		// external book references
		f = "SUM(A1:A3)+SUM(SheetA!A1:A3)+SUM([Book2]SheetA!A1:A3)+SUM([Book2]SheetB!A1:A3)";
		testFormulaMove(engine, sheetA, f, new SheetRegion(sheetA, "A1:A3"), 2, 0,
				"SUM(A3:A5)+SUM(SheetA!A3:A5)+SUM([Book2]SheetA!A1:A3)+SUM([Book2]SheetB!A1:A3)"); // shift in current book
		testFormulaMove(engine, sheetA, f, new SheetRegion(book2SheetA, "A1:A3"), 2, 0,
				"SUM(A1:A3)+SUM(SheetA!A1:A3)+SUM([Book2]SheetA!A3:A5)+SUM([Book2]SheetB!A1:A3)"); // shift in external book

		// absolute formula only effect copy and auto-fill operation
		// move, insert and delete operations still effect absolute formulas  
		f = "SUM($C3:$E5)+SUM(C$3:E$5)+SUM($C$3:$E$5)";
		testFormulaMove(engine, sheetA, f, new SheetRegion(sheetA, "C3:E5"), 2, 2,
				"SUM($E5:$G7)+SUM(E$5:G$7)+SUM($E$5:$G$7)");
		
		// 3D reference, don't get any effected
		f = "SUM(Sheet1:Sheet3!A1)";
		testFormulaMove(engine, sheetA, f, new SheetRegion(sheetA, "A1:A1"), 2, 2, f);
		
		// intersection
		f = "SUM(C3:E5)";
		testFormulaMove(engine, sheetA, f, new SheetRegion(sheetA, "D3:E5"), 0, 1, "SUM(C3:F5)");
		
		// area direction
		testFormulaMove(engine, sheetA, "SUM(C3:E5)", new SheetRegion(sheetA, "C3:E5"), 2, 2, "SUM(E5:G7)");
		testFormulaMove(engine, sheetA, "SUM(E5:C3)", new SheetRegion(sheetA, "C3:E5"), 2, 2, "SUM(E5:G7)");
		testFormulaMove(engine, sheetA, "SUM(C5:E3)", new SheetRegion(sheetA, "C3:E5"), 2, 2, "SUM(E5:G7)");
		testFormulaMove(engine, sheetA, "SUM(E3:C5)", new SheetRegion(sheetA, "C3:E5"), 2, 2, "SUM(E5:G7)");
	}

	private void testFormulaMove(FormulaEngine engine, SSheet sheet, String f, SheetRegion region, int rowOffset, int colOffset, String expected) {
		FormulaParseContext context = new FormulaParseContext(sheet, null);
		FormulaExpression expr = engine.move(f, region, rowOffset, colOffset, context);
		Assert.assertFalse(expr.hasError());
		Assert.assertEquals(expected, expr.getFormulaString());
	}
	
	@Test
	public void testFormulaShrink() {
		FormulaEngine engine = EngineFactory.getInstance().createFormulaEngine();
		SBook book1 = SBooks.createBook("Book1");
		SSheet sheetA = book1.createSheet("SheetA");
		SSheet sheetB = book1.createSheet("SheetB");
		SBook book2 = SBooks.createBook("Book2");
		SSheet book2SheetA = book2.createSheet("SheetA");
		new BookSeriesBuilderImpl().buildBookSeries(book1, book2);

		// the formula contains 3 region in current sheet
		// delete region won't cover region 1, complete cover region 2, and partial cover region 3
		String f = "SUM(C3:E5)+SUM(G3:I5)+SUM(K3:M5)";

		// delete cells and shift up
		boolean horizontal = false;
		// source region at top
		testFormulaShrink(f,sheetA, "G1:L1", null, horizontal, "SUM(C3:E5)+SUM(G2:I4)+SUM(K3:M5)", engine);
		testFormulaShrink(f,sheetA, "G1:L2", null, horizontal, "SUM(C3:E5)+SUM(G1:I3)+SUM(K3:M5)", engine);
		// source region overlapped
		testFormulaShrink(f,sheetA, "G3:L3", null, horizontal, "SUM(C3:E5)+SUM(G3:I4)+SUM(K3:M5)", engine); // 1 row
		testFormulaShrink(f,sheetA, "G4:L4", null, horizontal, "SUM(C3:E5)+SUM(G3:I4)+SUM(K3:M5)", engine); // 1 row
		testFormulaShrink(f,sheetA, "G5:L5", null, horizontal, "SUM(C3:E5)+SUM(G3:I4)+SUM(K3:M5)", engine); // 1 row
		testFormulaShrink(f,sheetA, "G2:L3", null, horizontal, "SUM(C3:E5)+SUM(G2:I3)+SUM(K3:M5)", engine); // 2 rows
		testFormulaShrink(f,sheetA, "G2:L4", null, horizontal, "SUM(C3:E5)+SUM(G2:I2)+SUM(K3:M5)", engine); // 2 rows
		testFormulaShrink(f,sheetA, "G4:L5", null, horizontal, "SUM(C3:E5)+SUM(G3:I3)+SUM(K3:M5)", engine); // 2 rows
		testFormulaShrink(f,sheetA, "G3:L5", null, horizontal, "SUM(C3:E5)+SUM(#REF!)+SUM(K3:M5)", engine); // 3 rows
//		testFormulaShrink(f,"G3:L5", horizontal, "SUM(C3:E5)+SUM(#REF!)+SUM(M3:M5)", engine, sheetA); // it's Excel approach 
		// source region at bottom
		testFormulaShrink(f,sheetA, "G6:L6", null, horizontal, f, engine);
		// external sheet with external book
		f = "SUM(A1:A3)+SUM(SheetA!A1:A3)+SUM(SheetB!A1:A3)+SUM([Book2]SheetA!A1:A3)";
		testFormulaShrink(f,sheetA, "A2:A2", null, horizontal, "SUM(A1:A2)+SUM(SheetA!A1:A2)+SUM(SheetB!A1:A3)+SUM([Book2]SheetA!A1:A3)", engine); // on SheetA
		testFormulaShrink(f,sheetB, "A2:A2", null, horizontal, "SUM(A1:A2)+SUM(SheetA!A1:A3)+SUM(SheetB!A1:A2)+SUM([Book2]SheetA!A1:A3)", engine); // on SheetB
		testFormulaShrink(f,sheetA, "A2:A2", book2SheetA, horizontal, "SUM(A1:A3)+SUM(SheetA!A1:A3)+SUM(SheetB!A1:A3)+SUM([Book2]SheetA!A1:A2)", engine); // on Book2 SheetA

		// delete cells and shift left
		f = "SUM(C3:E5)+SUM(C7:E9)+SUM(C11:E13)";
		horizontal = true;
		// source region at left
		testFormulaShrink(f,sheetA, "A7:A12", null, horizontal, "SUM(C3:E5)+SUM(B7:D9)+SUM(C11:E13)", engine);
		testFormulaShrink(f,sheetA, "A7:B12", null, horizontal, "SUM(C3:E5)+SUM(A7:C9)+SUM(C11:E13)", engine);
		// source region overlapped
		testFormulaShrink(f,sheetA, "C7:C12", null, horizontal, "SUM(C3:E5)+SUM(C7:D9)+SUM(C11:E13)", engine); // 1 column
		testFormulaShrink(f,sheetA, "D7:D12", null, horizontal, "SUM(C3:E5)+SUM(C7:D9)+SUM(C11:E13)", engine); // 1 column
		testFormulaShrink(f,sheetA, "E7:E12", null, horizontal, "SUM(C3:E5)+SUM(C7:D9)+SUM(C11:E13)", engine); // 1 column
		testFormulaShrink(f,sheetA, "B7:C12", null, horizontal, "SUM(C3:E5)+SUM(B7:C9)+SUM(C11:E13)", engine); // 2 columns
		testFormulaShrink(f,sheetA, "B7:D12", null, horizontal, "SUM(C3:E5)+SUM(B7:B9)+SUM(C11:E13)", engine); // 2 columns
		testFormulaShrink(f,sheetA, "D7:E12", null, horizontal, "SUM(C3:E5)+SUM(C7:C9)+SUM(C11:E13)", engine); // 2 columns
		testFormulaShrink(f,sheetA, "C7:E12", null, horizontal, "SUM(C3:E5)+SUM(#REF!)+SUM(C11:E13)", engine); // 3 columns
//		testFormulaShrink(f,"C7:E12", horizontal, "SUM(C3:E5)+SUM(#REF!)+SUM(C13:E13)", engine, sheetA); // it's Excel approach 
		// source region at right
		testFormulaShrink(f,sheetA, "F7:F12", null, horizontal, f, engine);
		// external sheet with external book
		f = "SUM(A1:C1)+SUM(SheetA!A1:C1)+SUM(SheetB!A1:C1)+SUM([Book2]SheetA!A1:C1)";
		testFormulaShrink(f,sheetA, "B1:B1", null, horizontal, "SUM(A1:B1)+SUM(SheetA!A1:B1)+SUM(SheetB!A1:C1)+SUM([Book2]SheetA!A1:C1)", engine); // on SheetA
		testFormulaShrink(f,sheetB, "B1:B1", null, horizontal, "SUM(A1:B1)+SUM(SheetA!A1:C1)+SUM(SheetB!A1:B1)+SUM([Book2]SheetA!A1:C1)", engine); // on SheetB
		testFormulaShrink(f,sheetA, "B1:B1", book2SheetA, horizontal, "SUM(A1:C1)+SUM(SheetA!A1:C1)+SUM(SheetB!A1:C1)+SUM([Book2]SheetA!A1:B1)", engine); // on Book2 SheetA
		
		// area direction
		testFormulaShrink("SUM(G3:I5)", sheetA, "G1:L1", null, false, "SUM(G2:I4)", engine);
		testFormulaShrink("SUM(G3:I5)", sheetA, "G1:L1", null, false, "SUM(G2:I4)", engine);
		testFormulaShrink("SUM(G3:I5)", sheetA, "G1:L1", null, false, "SUM(G2:I4)", engine);
		testFormulaShrink("SUM(G3:I5)", sheetA, "G1:L1", null, false, "SUM(G2:I4)", engine);
	}

	private void testFormulaShrink(String formula, SSheet formulaSheet, String region, SSheet regionSheet, boolean hrizontal, String expected, FormulaEngine engine) {
		if(regionSheet == null) {
			regionSheet = formulaSheet;
		}
		FormulaParseContext context = new FormulaParseContext(formulaSheet, null);
		FormulaExpression expr = engine.shrink(formula, new SheetRegion(regionSheet, region), hrizontal, context);
		Assert.assertFalse(expr.hasError());
		Assert.assertEquals(expected, expr.getFormulaString());
	}
	
	@Test
	public void testFormulaExtend() {
		FormulaEngine engine = EngineFactory.getInstance().createFormulaEngine();
		SBook book1 = SBooks.createBook("Book1");
		SSheet sheetA = book1.createSheet("SheetA");
		SSheet sheetB = book1.createSheet("SheetB");
		SBook book2 = SBooks.createBook("Book2");
		SSheet book2SheetA = book2.createSheet("SheetA");
		new BookSeriesBuilderImpl().buildBookSeries(book1, book2);

		// the formula contains 3 region in current sheet
		// target region won't cover region 1, complete cover region 2, and partial cover region 3
		String f = "SUM(C3:E5)+SUM(G3:I5)+SUM(K3:M5)";
		
		// delete cells and shift up
		boolean horizontal = false;
		// source region at top
		testFormulaExtend(f,sheetA, "G1:L1", null, horizontal, "SUM(C3:E5)+SUM(G4:I6)+SUM(K3:M5)", engine);
		testFormulaExtend(f,sheetA, "G1:L2", null, horizontal, "SUM(C3:E5)+SUM(G5:I7)+SUM(K3:M5)", engine);
		// source region overlapped
		testFormulaExtend(f,sheetA, "G3:L3", null, horizontal, "SUM(C3:E5)+SUM(G4:I6)+SUM(K3:M5)", engine); // 1 row
		testFormulaExtend(f,sheetA, "G4:L4", null, horizontal, "SUM(C3:E5)+SUM(G3:I6)+SUM(K3:M5)", engine); // 1 row
		testFormulaExtend(f,sheetA, "G5:L5", null, horizontal, "SUM(C3:E5)+SUM(G3:I6)+SUM(K3:M5)", engine); // 1 row
		testFormulaExtend(f,sheetA, "G2:L3", null, horizontal, "SUM(C3:E5)+SUM(G5:I7)+SUM(K3:M5)", engine); // 2 rows
		testFormulaExtend(f,sheetA, "G2:L4", null, horizontal, "SUM(C3:E5)+SUM(G6:I8)+SUM(K3:M5)", engine); // 2 rows
		testFormulaExtend(f,sheetA, "G4:L5", null, horizontal, "SUM(C3:E5)+SUM(G3:I7)+SUM(K3:M5)", engine); // 2 rows
		testFormulaExtend(f,sheetA, "G3:L5", null, horizontal, "SUM(C3:E5)+SUM(G6:I8)+SUM(K3:M5)", engine); // 3 rows
		// source region at bottom
		testFormulaExtend(f,sheetA, "G6:L6", null, horizontal, f, engine);
		// external sheet with external book
		f = "SUM(A1:A3)+SUM(SheetA!A1:A3)+SUM(SheetB!A1:A3)+SUM([Book2]SheetA!A1:A3)";
		testFormulaExtend(f,sheetA, "A2:A2", null, horizontal, "SUM(A1:A4)+SUM(SheetA!A1:A4)+SUM(SheetB!A1:A3)+SUM([Book2]SheetA!A1:A3)", engine); // on SheetA
		testFormulaExtend(f,sheetB, "A2:A2", null, horizontal, "SUM(A1:A4)+SUM(SheetA!A1:A3)+SUM(SheetB!A1:A4)+SUM([Book2]SheetA!A1:A3)", engine); // on SheetB
		testFormulaExtend(f,sheetB, "A2:A2", book2SheetA, horizontal, "SUM(A1:A3)+SUM(SheetA!A1:A3)+SUM(SheetB!A1:A3)+SUM([Book2]SheetA!A1:A4)", engine); // on Book2 SheetA
		
		// delete cells and shift left
		f = "SUM(C3:E5)+SUM(C7:E9)+SUM(C11:E13)";
		horizontal = true;
		// source region at left
		testFormulaExtend(f,sheetA, "A7:A12", null, horizontal, "SUM(C3:E5)+SUM(D7:F9)+SUM(C11:E13)", engine);
		testFormulaExtend(f,sheetA, "A7:B12", null, horizontal, "SUM(C3:E5)+SUM(E7:G9)+SUM(C11:E13)", engine);
		// source region overlapped
		testFormulaExtend(f,sheetA, "C7:C12", null, horizontal, "SUM(C3:E5)+SUM(D7:F9)+SUM(C11:E13)", engine); // 1 column
		testFormulaExtend(f,sheetA, "D7:D12", null, horizontal, "SUM(C3:E5)+SUM(C7:F9)+SUM(C11:E13)", engine); // 1 column
		testFormulaExtend(f,sheetA, "E7:E12", null, horizontal, "SUM(C3:E5)+SUM(C7:F9)+SUM(C11:E13)", engine); // 1 column
		testFormulaExtend(f,sheetA, "B7:C12", null, horizontal, "SUM(C3:E5)+SUM(E7:G9)+SUM(C11:E13)", engine); // 2 columns
		testFormulaExtend(f,sheetA, "B7:D12", null, horizontal, "SUM(C3:E5)+SUM(F7:H9)+SUM(C11:E13)", engine); // 2 columns
		testFormulaExtend(f,sheetA, "D7:E12", null, horizontal, "SUM(C3:E5)+SUM(C7:G9)+SUM(C11:E13)", engine); // 2 columns
		testFormulaExtend(f,sheetA, "C7:E12", null, horizontal, "SUM(C3:E5)+SUM(F7:H9)+SUM(C11:E13)", engine); // 3 columns
		// source region at right
		testFormulaExtend(f,sheetA, "F7:F12", null, horizontal, f, engine);
		// external sheet with external book
		f = "SUM(A1:C1)+SUM(SheetA!A1:C1)+SUM(SheetB!A1:C1)+SUM([Book2]SheetA!A1:C1)";
		testFormulaExtend(f,sheetA, "B1:B1", null, horizontal, "SUM(A1:D1)+SUM(SheetA!A1:D1)+SUM(SheetB!A1:C1)+SUM([Book2]SheetA!A1:C1)", engine); // on SheetA
		testFormulaExtend(f,sheetB, "B1:B1", null, horizontal, "SUM(A1:D1)+SUM(SheetA!A1:C1)+SUM(SheetB!A1:D1)+SUM([Book2]SheetA!A1:C1)", engine); // on SheetB
		testFormulaExtend(f,sheetA, "B1:B1", book2SheetA, horizontal, "SUM(A1:C1)+SUM(SheetA!A1:C1)+SUM(SheetB!A1:C1)+SUM([Book2]SheetA!A1:D1)", engine); // on Book2 SheetA
		
		// area direction
		testFormulaExtend("SUM(G3:I5)", sheetA, "G1:L1", null, false, "SUM(G4:I6)", engine);
		testFormulaExtend("SUM(I5:G3)", sheetA, "G1:L1", null, false, "SUM(G4:I6)", engine);
		testFormulaExtend("SUM(G5:I3)", sheetA, "G1:L1", null, false, "SUM(G4:I6)", engine);
		testFormulaExtend("SUM(I3:G5)", sheetA, "G1:L1", null, false, "SUM(G4:I6)", engine);
	}
	
	private void testFormulaExtend(String formula, SSheet formulaSheet, String region, SSheet regionSheet, boolean hrizontal, String expected, FormulaEngine engine) {
		if(regionSheet == null) {
			regionSheet = formulaSheet;
		}
		FormulaParseContext context = new FormulaParseContext(formulaSheet, null);
		FormulaExpression expr = engine.extend(formula, new SheetRegion(regionSheet, region), hrizontal, context);
		Assert.assertFalse(expr.hasError());
		Assert.assertEquals(expected, expr.getFormulaString());
	}
	
	@Test
	public void testFormulaShift() {
		FormulaEngine engine = EngineFactory.getInstance().createFormulaEngine();
		SBook book1 = SBooks.createBook("Book1");
		SSheet sheetA = book1.createSheet("SheetA");
		
		String f = "SUM(B2,B3:C4,Sheet1!B5:D7,Sheet2:Sheet3!B8:E11,[Book2.xlsx]Sheet1!E11:G13)";
		// no shift
		testFormulaShift(engine, sheetA, f, 0, 0, f);
		// row
		testFormulaShift(engine, sheetA, f, -1, 0, "SUM(B1,B2:C3,Sheet1!B4:D6,Sheet2:Sheet3!B7:E10,[Book2.xlsx]Sheet1!E10:G12)");
		testFormulaShift(engine, sheetA, f, 1, 0, "SUM(B3,B4:C5,Sheet1!B6:D8,Sheet2:Sheet3!B9:E12,[Book2.xlsx]Sheet1!E12:G14)");
		testFormulaShift(engine, sheetA, f, 2, 0, "SUM(B4,B5:C6,Sheet1!B7:D9,Sheet2:Sheet3!B10:E13,[Book2.xlsx]Sheet1!E13:G15)");
		// column
		testFormulaShift(engine, sheetA, f, 0, -1, "SUM(A2,A3:B4,Sheet1!A5:C7,Sheet2:Sheet3!A8:D11,[Book2.xlsx]Sheet1!D11:F13)");
		testFormulaShift(engine, sheetA, f, 0, 1, "SUM(C2,C3:D4,Sheet1!C5:E7,Sheet2:Sheet3!C8:F11,[Book2.xlsx]Sheet1!F11:H13)");
		testFormulaShift(engine, sheetA, f, 0, 2, "SUM(D2,D3:E4,Sheet1!D5:F7,Sheet2:Sheet3!D8:G11,[Book2.xlsx]Sheet1!G11:I13)");
		// both
		testFormulaShift(engine, sheetA, f, -1, -1, "SUM(A1,A2:B3,Sheet1!A4:C6,Sheet2:Sheet3!A7:D10,[Book2.xlsx]Sheet1!D10:F12)");
		testFormulaShift(engine, sheetA, f, 1, 2, "SUM(D3,D4:E5,Sheet1!D6:F8,Sheet2:Sheet3!D9:G12,[Book2.xlsx]Sheet1!G12:I14)");
		testFormulaShift(engine, sheetA, f, 2, 1, "SUM(C4,C5:D6,Sheet1!C7:E9,Sheet2:Sheet3!C10:F13,[Book2.xlsx]Sheet1!F13:H15)");
		// out of bounds
		testFormulaShift(engine, sheetA, f, -2, 0, "SUM(#REF!,B1:C2,Sheet1!B3:D5,Sheet2:Sheet3!B6:E9,[Book2.xlsx]Sheet1!E9:G11)");
		testFormulaShift(engine, sheetA, f, 0, -2, "SUM(#REF!,#REF!,Sheet1!#REF!,Sheet2:Sheet3!#REF!,[Book2.xlsx]Sheet1!C11:E13)");
		testFormulaShift(engine, sheetA, f, -2, -2, "SUM(#REF!,#REF!,Sheet1!#REF!,Sheet2:Sheet3!#REF!,[Book2.xlsx]Sheet1!C9:E11)");
		testFormulaShift(engine, sheetA, f, book1.getMaxRowSize() - 10, 0, "SUM(B1048568,B1048569:C1048570,Sheet1!B1048571:D1048573,Sheet2:Sheet3!#REF!,[Book2.xlsx]Sheet1!#REF!)");
		testFormulaShift(engine, sheetA, f, 0, book1.getMaxColumnSize() - 4, "SUM(XFB2,XFB3:XFC4,Sheet1!XFB5:XFD7,Sheet2:Sheet3!#REF!,[Book2.xlsx]Sheet1!#REF!)");
		
		// absolute
		testFormulaShift(engine, sheetA, "SUM(B2,B$2,$B2,$B$2)", 1, 1, "SUM(C3,C$2,$B3,$B$2)"); // ref
		testFormulaShift(engine, sheetA, "SUM(B2:C3,$B2:C3,B$2:C3,B2:$C3,B2:C$3,$B$2:C3,B2:$C$3,$B2:$C3,B$2:C$3,B$2:$C3,$B$2:$C3,$B$2:C$3,$B2:$C$3,B$2:$C$3,$B$2:$C$3)",
				1, 1, "SUM(C3:D4,$B3:D4,C$2:D4,C3:$C4,C3:D$3,$B$2:D4,C3:$C$3,$B3:$C4,C$2:D$3,C$2:$C4,$B$2:$C4,$B$2:D$3,$B3:$C$3,C$2:$C$3,$B$2:$C$3)"); // area
		testFormulaShift(engine, sheetA, "SUM($B2,B$3:C4,Sheet1!B5:$D7,Sheet2!B8:E$11,[Book2.xlsx]Sheet1!$E11:G$13)", 
				1, 1, "SUM($B3,C$3:D5,Sheet1!C6:$D8,Sheet2!C9:F$11,[Book2.xlsx]Sheet1!$E12:H$13)");
		
		// area direction
		testFormulaShift(engine, sheetA, "SUM(G6:H11)", 2, 3, "SUM(J8:K13)");
		testFormulaShift(engine, sheetA, "SUM(H11:G6)", 2, 3, "SUM(J8:K13)");
		testFormulaShift(engine, sheetA, "SUM(G11:H6)", 2, 3, "SUM(J8:K13)");
		testFormulaShift(engine, sheetA, "SUM(H6:G11)", 2, 3, "SUM(J8:K13)");
	}
	
	private void testFormulaShift(FormulaEngine engine, SSheet sheet, String formula, int rowOffset, int columnOffset, String expected) {
		FormulaParseContext context = new FormulaParseContext(sheet, null);
		FormulaExpression expr = engine.shift(formula, rowOffset, columnOffset, context);
		Assert.assertFalse(expr.hasError());
		Assert.assertEquals(expected, expr.getFormulaString());
	}
	
	@Test
	public void testFormulaTranspose() {
		FormulaEngine engine = EngineFactory.getInstance().createFormulaEngine();
		SBook book1 = SBooks.createBook("Book1");
		SSheet sheetA = book1.createSheet("SheetA");

		// normal
		String f = "SUM(G6,Sheet1!G6,Sheet2:Sheet3!G6:I7,[Book2.xlsx]Sheet1!G6:H11)";
		String C5 = testFormulaTranspose(engine, sheetA, f, "C5", "SUM(D9,Sheet1!D9,Sheet2:Sheet3!D9:E11,[Book2.xlsx]Sheet1!D9:I10)"); 
		String D5 = testFormulaTranspose(engine, sheetA, f, "D5", "SUM(E8,Sheet1!E8,Sheet2:Sheet3!E8:F10,[Book2.xlsx]Sheet1!E8:J9)"); 
		String C6 = testFormulaTranspose(engine, sheetA, f, "C6", "SUM(C10,Sheet1!C10,Sheet2:Sheet3!C10:D12,[Book2.xlsx]Sheet1!C10:H11)");
		String D6 = testFormulaTranspose(engine, sheetA, f, "D6", "SUM(D9,Sheet1!D9,Sheet2:Sheet3!D9:E11,[Book2.xlsx]Sheet1!D9:I10)");
		// different origin positions
		testFormulaTranspose(engine, sheetA, f, "H2", "SUM(L1,Sheet1!L1,Sheet2:Sheet3!L1:M3,[Book2.xlsx]Sheet1!L1:Q2)"); 
		testFormulaTranspose(engine, sheetA, f, "K8", "SUM(I4,Sheet1!I4,Sheet2:Sheet3!I4:J6,[Book2.xlsx]Sheet1!I4:N5)");
		testFormulaTranspose(engine, sheetA, f, "F11", "SUM(A12,Sheet1!A12,Sheet2:Sheet3!A12:B14,[Book2.xlsx]Sheet1!A12:F13)");
		// absolute
		testFormulaTranspose(engine, sheetA, "SUM(G12,G$12,$G12,$G$12)", "H2", "SUM(R1,G$12,$G12,$G$12)"); 
		testFormulaTranspose(engine, sheetA, "SUM(H9:J15,$H9:J15,H$9:J15,H9:$J15,H9:J$15,$H$9:J15,$H9:$J15,$H9:J$15,H$9:$J15,H$9:J$15,H9:$J$15,$H$9:$J15,$H$9:J$15,$H9:$J$15,H$9:$J$15,$H$9:$J$15)",
				"K8", "SUM(L5:R7,L$5:R7,$L5:R7,L5:R$7,L5:$R7,$L$5:R7,$H9:$J15,$H9:J$15,H$9:$J15,H$9:J$15,L5:$R$7,$H$9:$J15,$H$9:J$15,$H9:$J$15,H$9:$J$15,$H$9:$J$15)");
		testFormulaTranspose(engine, sheetA, "SUM($G9,G$9:J15,Sheet1!G10:$J16,Sheet2:Sheet3!H11:L$12,[Book2.xlsx]Sheet1!$G9:O$10)",
				"F11", "SUM($G9,$D12:J15,Sheet1!E12:K$15,Sheet2:Sheet3!F13:$G17,[Book2.xlsx]Sheet1!$G9:O$10)");

		// transpose than shift
		testFormulaShift(engine, sheetA, D5, 1, -1, C5);
		testFormulaShift(engine, sheetA, C6, -1, 1, C5);
		testFormulaShift(engine, sheetA, D6, 0, 0, C5);
		
		// area direction
		testFormulaTranspose(engine, sheetA, "SUM(G6:H11)", "C5", "SUM(D9:I10)");
		testFormulaTranspose(engine, sheetA, "SUM(H11:G6)", "C5", "SUM(D9:I10)");
		testFormulaTranspose(engine, sheetA, "SUM(G11:H6)", "C5", "SUM(D9:I10)");
		testFormulaTranspose(engine, sheetA, "SUM(H6:G11)", "C5", "SUM(D9:I10)");
	}
	
	private String testFormulaTranspose(FormulaEngine engine, SSheet sheet, String formula, String origin, String expected) {
		SheetRegion o = new SheetRegion(sheet, origin);
		FormulaParseContext context = new FormulaParseContext(sheet, null);
		FormulaExpression expr = engine.transpose(formula, o.getRow(), o.getColumn(), context);
		Assert.assertFalse(expr.hasError());
		String actual = expr.getFormulaString();
		Assert.assertEquals(expected, actual);
		return actual;
	}
	
	@Test
	public void testFormulaRenameSheet() {
		FormulaEngine engine = EngineFactory.getInstance().createFormulaEngine();
		SBook bookA = SBooks.createBook("BookA");
		SSheet sheetX = bookA.createSheet("SheetX");
		SSheet sheet1 = bookA.createSheet("Sheet1");
		SSheet sheet2 = bookA.createSheet("Sheet2");
		bookA.createSheet("Sheet3");
		bookA.createSheet("Sheet4");
		SSheet sheet5 = bookA.createSheet("Sheet5");
		bookA.createSheet("Sheet6");
		SBook bookB = SBooks.createBook("BookB.xlsx");
		bookB.createSheet("Sheet1");
		new BookSeriesBuilderImpl().buildBookSeries(bookA, bookB);

		String f = "SUM(Sheet1!A1,Sheet2:Sheet5!A1,Sheet6!A1,[BookB.xlsx]Sheet1!A1)";

		// normal
		testFormulaRenameSheet(engine, sheetX, f, bookA, "Sheet1", "sht1", "SUM(sht1!A1,Sheet2:Sheet5!A1,Sheet6!A1,[BookB.xlsx]Sheet1!A1)");
		testFormulaRenameSheet(engine, sheetX, f, bookA, "Sheet2", "sht2", "SUM(Sheet1!A1,sht2:Sheet5!A1,Sheet6!A1,[BookB.xlsx]Sheet1!A1)");
		testFormulaRenameSheet(engine, sheetX, f, bookA, "Sheet3", "sht3", f);
		testFormulaRenameSheet(engine, sheetX, f, bookA, "Sheet4", "sht4", f);
		testFormulaRenameSheet(engine, sheetX, f, bookA, "Sheet5", "sht5", "SUM(Sheet1!A1,Sheet2:sht5!A1,Sheet6!A1,[BookB.xlsx]Sheet1!A1)");
		testFormulaRenameSheet(engine, sheetX, f, bookA, "Sheet6", "sht6", "SUM(Sheet1!A1,Sheet2:Sheet5!A1,sht6!A1,[BookB.xlsx]Sheet1!A1)");

		// duplicate name
		// modification should be successful but just fail to eval.
		testFormulaRenameSheet(engine, sheetX, f, bookA, "Sheet1", "Sheet2", "SUM(Sheet2!A1,Sheet2:Sheet5!A1,Sheet6!A1,[BookB.xlsx]Sheet1!A1)");
		testFormulaRenameSheet(engine, sheetX, f, bookA, "Sheet2", "Sheet5", "SUM(Sheet1!A1,Sheet5!A1,Sheet6!A1,[BookB.xlsx]Sheet1!A1)"); // merge sheet range

		// external book
		testFormulaRenameSheet(engine, sheetX, f, bookB, "Sheet1", "sht1", "SUM(Sheet1!A1,Sheet2:Sheet5!A1,Sheet6!A1,[BookB.xlsx]sht1!A1)");
		testFormulaRenameSheet(engine, sheetX, f, bookB, "Sheet2", "sht2", f);
		
		// delete sheet
		// this spec. isn't compatible to Excel, Excel approach is more smart.
		f = "SUM(A1,Sheet1!A1,Sheet2:Sheet5!A1,Sheet6!A1,[BookB.xlsx]Sheet1!A1)";
		testFormulaRenameSheet(engine, sheet1, f, bookA, "Sheet1", null, "SUM(A1,'#REF'!A1,Sheet2:Sheet5!A1,Sheet6!A1,[BookB.xlsx]Sheet1!A1)");
		testFormulaRenameSheet(engine, sheet2, f, bookA, "Sheet2", null, "SUM(A1,Sheet1!A1,'#REF'!A1,Sheet6!A1,[BookB.xlsx]Sheet1!A1)");
		testFormulaRenameSheet(engine, sheet5, f, bookA, "Sheet5", null, "SUM(A1,Sheet1!A1,'#REF'!A1,Sheet6!A1,[BookB.xlsx]Sheet1!A1)");
		testFormulaRenameSheet(engine, sheet1, f, bookB, "Sheet1", null, "SUM(A1,Sheet1!A1,Sheet2:Sheet5!A1,Sheet6!A1,'[BookB.xlsx]#REF'!A1)");
	}
	
	private void testFormulaRenameSheet(FormulaEngine engine, SSheet formulaSheet, String formula, SBook targetBook, String oldSheetName, String newSheetName, String expected) {
		FormulaParseContext context = new FormulaParseContext(formulaSheet, null);
		FormulaExpression expr = engine.renameSheet(formula, targetBook, oldSheetName, newSheetName, context);
		Assert.assertFalse(expr.hasError());
		Assert.assertEquals(expected, expr.getFormulaString());
	}
}
