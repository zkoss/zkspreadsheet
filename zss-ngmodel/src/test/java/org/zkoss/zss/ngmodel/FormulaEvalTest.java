/* FormulaEvalTest.java

	Purpose:
		
	Description:
		
	History:
		Dec 9, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.ngmodel;

import java.io.Closeable;
import java.io.InputStream;
import java.text.MessageFormat;
import java.util.Locale;
import java.util.Set;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.util.Locales;
import org.zkoss.zss.ngapi.NImporter;
import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngapi.NRanges;
import org.zkoss.zss.ngapi.impl.imexp.ExcelImportFactory;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.impl.AbstractBookSeriesAdv;
import org.zkoss.zss.ngmodel.impl.BookSeriesBuilderImpl;
import org.zkoss.zss.ngmodel.impl.NameRefImpl;
import org.zkoss.zss.ngmodel.impl.RefImpl;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.sys.formula.FormulaParseContext;

/**
 * @author Pao
 */
public class FormulaEvalTest {

	private NImporter importer;

	@Before
	public void beforeTest() {
		importer= new ExcelImportFactory().createImporter();
		Locales.setThreadLocal(Locale.TAIWAN);
	}

	@Test
	public void testBasicEvaluation() {
		NBook book = NBooks.createBook("book1");
		NSheet sheet1 = book.createSheet("Sheet1");
		sheet1.getCell(0, 0).setFormulaValue("SUM(C1:C10)");
		sheet1.getCell(0, 1).setNumberValue(55.0);
		sheet1.getCell(1, 0).setFormulaValue("AVERAGE(Sheet2!C1:C10)");
		sheet1.getCell(1, 1).setNumberValue(5.5);
		for(int r = 0; r < 10; ++r) {
			sheet1.getCell(r, 2).setValue(r + 1);
		}
		NSheet sheet2 = book.createSheet("Sheet2");
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
		NBook book = NBooks.createBook("book1");
		NSheet sheet1 = book.createSheet("Sheet1");
		NSheet sheet2 = book.createSheet("Sheet2");
		NSheet sheetA = book.createSheet("SheetA");

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
			NCell cell = sheet1.getCell(i, 0);
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

		NBook book = NBooks.createBook("book1");
		NSheet sheet1 = book.createSheet("Sheet1");
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

		NBook book1 = NBooks.createBook("book1");
		NSheet sheetA = book1.createSheet("SheetA");
		NBook book2 = NBooks.createBook("book2");
		NSheet sheetB = book2.createSheet("SheetB");
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
	
	private NCell getCell(NSheet sheet, String ref) {
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
		NBook book = NBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		DependencyTable table = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
		
		NSheet sheet1 = book.createSheet("Sheet1");
		NSheet sheet2 = book.createSheet("Sheet2");

		NCell cell = sheet1.getCell(0, 0);
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
		NRange r = NRanges.range(sheet2, 1, 0);
		
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
		NBook book = NBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		DependencyTable table = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
		
		NSheet sheet1 = book.createSheet("Sheet1");

		NCell cell = sheet1.getCell(0, 0);
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
		
		
		NSheet sheet2 = book.createSheet("Sheet2");
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
		NBook book = NBooks.createBook("book1");
		NSheet sheet1 = book.createSheet("Sheet1");
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

	public NBook getBook(String path, String bookName) {
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
		NBook book = getBook("book/formula-eval.xlsx", "Book1");
		Assert.assertNotNull(book);
		NSheet sheet = book.getSheetByName("array");
		Assert.assertNotNull(sheet);

		// check formula
		int c = 2;
		for(int r = 1; r <= 6; ++r) { // C2:C7
			NCell cell = sheet.getCell(r, c);
			Assert.assertEquals("A2:A6", cell.getFormulaValue());
		}

		// check value
		c = 0;
		for(int r = 1; r <= 5; ++r) { // A2:A6
			NCell cell = sheet.getCell(r, c);
			Assert.assertEquals(r, cell.getNumberValue().intValue());
		}
		c = 2;
		for(int r = 1; r <= 5; ++r) { // C2:C6
			NCell cell = sheet.getCell(r, c);
			Assert.assertEquals(r, cell.getNumberValue().intValue());
		}

		// check C7
		NCell C7 = sheet.getCell(6, 2);
		Assert.assertEquals(CellType.ERROR, C7.getFormulaResultType());
		Assert.assertEquals(ErrorValue.INVALID_VALUE, C7.getErrorValue().getCode());
	}
	
	@Test
	@Ignore
	public void testArrayFormula() {
		NBook book = getBook("book/formula-eval.xlsx", "Book1");
		NSheet sheet = book.getSheetByName("array");
		Assert.assertNotNull(book);
		Assert.assertNotNull(sheet);

		// check eval. value
		NCell e2 = sheet.getCell(1, 4);
		Assert.assertFalse(e2.isNull());
		Assert.assertEquals(CellType.NUMBER, e2.getFormulaResultType());
		Assert.assertEquals(11, e2.getNumberValue().intValue());
	}
	
	@Test
	public void testExternalBookReference() {
		NBook book1, book2, book3;
		NSheet sheet1, sheet2;
		
		// direct creation
		book1 = NBooks.createBook("Book1");
		book2 = NBooks.createBook("Book2");
		book1.createSheet("external").getCell(0, 1).setFormulaValue("SUM(3 + [Book2]ref!C6 - 3)"); // B1
		book2.createSheet("ref").getCell(5, 2).setNumberValue(5.0);	// C6
		testExternalBookReference(book1, book2);
		
		// direct creation - complex 
		book1 = NBooks.createBook("Book1");
		book2 = NBooks.createBook("Book2");
		book3 = NBooks.createBook("Book3");
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

	public void testExternalBookReference(NBook... books) {
		Assert.assertTrue(books.length >= 2);
		NBook book1 = books[0];

		// update book series
		for(NBook book : books) {
			Assert.assertNotNull(book);
		}
		new BookSeriesBuilderImpl().buildBookSeries(books); 

		// check evaluated value of B1
		NSheet sheet = book1.getSheetByName("external");
		NCell b1 = sheet.getCell(0, 1);
		Assert.assertFalse(b1.isNull());
		Assert.assertEquals(CellType.NUMBER, b1.getFormulaResultType());
		Assert.assertEquals(5, b1.getNumberValue().intValue());
	}
	
	@Test
	public void testFormulaMove() {

		FormulaEngine engine = EngineFactory.getInstance().createFormulaEngine();
		NBook book1 = NBooks.createBook("Book1");
		NSheet sheetA = book1.createSheet("SheetA");
		NSheet sheetB = book1.createSheet("SheetB");
		NBook book2 = NBooks.createBook("Book2");
		NSheet book2SheetA = book2.createSheet("SheetA");
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
		

	}

	private void testFormulaMove(FormulaEngine engine, NSheet sheet, String f, SheetRegion region, int rowOffset, int colOffset, String expected) {
		FormulaParseContext context = new FormulaParseContext(sheet, null);
		FormulaExpression expr = engine.move(f, region, rowOffset, colOffset, context);
		Assert.assertFalse(expr.hasError());
		Assert.assertEquals(expected, expr.getFormulaString());
	}
	
	@Test
	public void testFormulaShrink() {
		FormulaEngine engine = EngineFactory.getInstance().createFormulaEngine();
		NBook book1 = NBooks.createBook("Book1");
		NSheet sheetA = book1.createSheet("SheetA");

		// the formula contains 3 region in current sheet
		// delete region won't cover region 1, complete cover region 2, and partial cover region 3
		String f = "SUM(C3:E5)+SUM(G3:I5)+SUM(K3:M5)";

		// delete cells and shift up
		boolean horizontal = false;
		// source region at top
		testFormulaShrink(f,"G1:L1", horizontal, "SUM(C3:E5)+SUM(G2:I4)+SUM(K3:M5)", engine, sheetA);
		testFormulaShrink(f,"G1:L2", horizontal, "SUM(C3:E5)+SUM(G1:I3)+SUM(K3:M5)", engine, sheetA);
		// source region overlapped
		testFormulaShrink(f,"G3:L3", horizontal, "SUM(C3:E5)+SUM(G3:I4)+SUM(K3:M5)", engine, sheetA); // 1 row
		testFormulaShrink(f,"G4:L4", horizontal, "SUM(C3:E5)+SUM(G3:I4)+SUM(K3:M5)", engine, sheetA); // 1 row
		testFormulaShrink(f,"G5:L5", horizontal, "SUM(C3:E5)+SUM(G3:I4)+SUM(K3:M5)", engine, sheetA); // 1 row
		testFormulaShrink(f,"G2:L3", horizontal, "SUM(C3:E5)+SUM(G2:I3)+SUM(K3:M5)", engine, sheetA); // 2 rows
		testFormulaShrink(f,"G2:L4", horizontal, "SUM(C3:E5)+SUM(G2:I2)+SUM(K3:M5)", engine, sheetA); // 2 rows
		testFormulaShrink(f,"G4:L5", horizontal, "SUM(C3:E5)+SUM(G3:I3)+SUM(K3:M5)", engine, sheetA); // 2 rows
		testFormulaShrink(f,"G3:L5", horizontal, "SUM(C3:E5)+SUM(#REF!)+SUM(K3:M5)", engine, sheetA); // 3 rows
//		testFormulaShrink(f,"G3:L5", horizontal, "SUM(C3:E5)+SUM(#REF!)+SUM(M3:M5)", engine, sheetA); // it's Excel approach 
		// source region at bottom
		testFormulaShrink(f,"G6:L6", horizontal, f, engine, sheetA);

		// delete cells and shift left
		f = "SUM(C3:E5)+SUM(C7:E9)+SUM(C11:E13)";
		horizontal = true;
		// source region at left
		testFormulaShrink(f,"A7:A12", horizontal, "SUM(C3:E5)+SUM(B7:D9)+SUM(C11:E13)", engine, sheetA);
		testFormulaShrink(f,"A7:B12", horizontal, "SUM(C3:E5)+SUM(A7:C9)+SUM(C11:E13)", engine, sheetA);
		// source region overlapped
		testFormulaShrink(f,"C7:C12", horizontal, "SUM(C3:E5)+SUM(C7:D9)+SUM(C11:E13)", engine, sheetA); // 1 column
		testFormulaShrink(f,"D7:D12", horizontal, "SUM(C3:E5)+SUM(C7:D9)+SUM(C11:E13)", engine, sheetA); // 1 column
		testFormulaShrink(f,"E7:E12", horizontal, "SUM(C3:E5)+SUM(C7:D9)+SUM(C11:E13)", engine, sheetA); // 1 column
		testFormulaShrink(f,"B7:C12", horizontal, "SUM(C3:E5)+SUM(B7:C9)+SUM(C11:E13)", engine, sheetA); // 2 columns
		testFormulaShrink(f,"B7:D12", horizontal, "SUM(C3:E5)+SUM(B7:B9)+SUM(C11:E13)", engine, sheetA); // 2 columns
		testFormulaShrink(f,"D7:E12", horizontal, "SUM(C3:E5)+SUM(C7:C9)+SUM(C11:E13)", engine, sheetA); // 2 columns
		testFormulaShrink(f,"C7:E12", horizontal, "SUM(C3:E5)+SUM(#REF!)+SUM(C11:E13)", engine, sheetA); // 3 columns
//		testFormulaShrink(f,"C7:E12", horizontal, "SUM(C3:E5)+SUM(#REF!)+SUM(C13:E13)", engine, sheetA); // it's Excel approach 
		// source region at right
		testFormulaShrink(f,"F7:F12", horizontal, f, engine, sheetA);
	}

	private void testFormulaShrink(String formula, String region, boolean hrizontal, String expected, FormulaEngine engine, NSheet sheet) {
		FormulaParseContext context = new FormulaParseContext(sheet, null);
		FormulaExpression expr = engine.shrink(formula, new SheetRegion(sheet, region), hrizontal, context);
		Assert.assertFalse(expr.hasError());
		Assert.assertEquals(expected, expr.getFormulaString());
	}
	
	@Test
	public void testFormulaExtend() {
		FormulaEngine engine = EngineFactory.getInstance().createFormulaEngine();
		NBook book1 = NBooks.createBook("Book1");
		NSheet sheetA = book1.createSheet("SheetA");

		// the formula contains 3 region in current sheet
		// target region won't cover region 1, complete cover region 2, and partial cover region 3
		String f = "SUM(C3:E5)+SUM(G3:I5)+SUM(K3:M5)";
		
		// delete cells and shift up
		boolean horizontal = false;
		// source region at top
		testFormulaExtend(f,"G1:L1", horizontal, "SUM(C3:E5)+SUM(G4:I6)+SUM(K3:M5)", engine, sheetA);
		testFormulaExtend(f,"G1:L2", horizontal, "SUM(C3:E5)+SUM(G5:I7)+SUM(K3:M5)", engine, sheetA);
		// source region overlapped
		testFormulaExtend(f,"G3:L3", horizontal, "SUM(C3:E5)+SUM(G4:I6)+SUM(K3:M5)", engine, sheetA); // 1 row
		testFormulaExtend(f,"G4:L4", horizontal, "SUM(C3:E5)+SUM(G3:I6)+SUM(K3:M5)", engine, sheetA); // 1 row
		testFormulaExtend(f,"G5:L5", horizontal, "SUM(C3:E5)+SUM(G3:I6)+SUM(K3:M5)", engine, sheetA); // 1 row
		testFormulaExtend(f,"G2:L3", horizontal, "SUM(C3:E5)+SUM(G5:I7)+SUM(K3:M5)", engine, sheetA); // 2 rows
		testFormulaExtend(f,"G2:L4", horizontal, "SUM(C3:E5)+SUM(G6:I8)+SUM(K3:M5)", engine, sheetA); // 2 rows
		testFormulaExtend(f,"G4:L5", horizontal, "SUM(C3:E5)+SUM(G3:I7)+SUM(K3:M5)", engine, sheetA); // 2 rows
		testFormulaExtend(f,"G3:L5", horizontal, "SUM(C3:E5)+SUM(G6:I8)+SUM(K3:M5)", engine, sheetA); // 3 rows
		// source region at bottom
		testFormulaExtend(f,"G6:L6", horizontal, f, engine, sheetA);

		// delete cells and shift left
		f = "SUM(C3:E5)+SUM(C7:E9)+SUM(C11:E13)";
		horizontal = true;
		// source region at left
		testFormulaExtend(f,"A7:A12", horizontal, "SUM(C3:E5)+SUM(D7:F9)+SUM(C11:E13)", engine, sheetA);
		testFormulaExtend(f,"A7:B12", horizontal, "SUM(C3:E5)+SUM(E7:G9)+SUM(C11:E13)", engine, sheetA);
		// source region overlapped
		testFormulaExtend(f,"C7:C12", horizontal, "SUM(C3:E5)+SUM(D7:F9)+SUM(C11:E13)", engine, sheetA); // 1 column
		testFormulaExtend(f,"D7:D12", horizontal, "SUM(C3:E5)+SUM(C7:F9)+SUM(C11:E13)", engine, sheetA); // 1 column
		testFormulaExtend(f,"E7:E12", horizontal, "SUM(C3:E5)+SUM(C7:F9)+SUM(C11:E13)", engine, sheetA); // 1 column
		testFormulaExtend(f,"B7:C12", horizontal, "SUM(C3:E5)+SUM(E7:G9)+SUM(C11:E13)", engine, sheetA); // 2 columns
		testFormulaExtend(f,"B7:D12", horizontal, "SUM(C3:E5)+SUM(F7:H9)+SUM(C11:E13)", engine, sheetA); // 2 columns
		testFormulaExtend(f,"D7:E12", horizontal, "SUM(C3:E5)+SUM(C7:G9)+SUM(C11:E13)", engine, sheetA); // 2 columns
		testFormulaExtend(f,"C7:E12", horizontal, "SUM(C3:E5)+SUM(F7:H9)+SUM(C11:E13)", engine, sheetA); // 3 columns
		// source region at right
		testFormulaExtend(f,"F7:F12", horizontal, f, engine, sheetA);
	}
	
	private void testFormulaExtend(String formula, String region, boolean hrizontal, String expected, FormulaEngine engine, NSheet sheet) {
		FormulaParseContext context = new FormulaParseContext(sheet, null);
		FormulaExpression expr = engine.extend(formula, new SheetRegion(sheet, region), hrizontal, context);
		Assert.assertFalse(expr.hasError());
		Assert.assertEquals(expected, expr.getFormulaString());
	}
}
