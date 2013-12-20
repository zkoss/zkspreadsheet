/* FormulaEvalTest.java

	Purpose:
		
	Description:
		
	History:
		Dec 9, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.ngmodel;

import java.util.Locale;
import java.util.Set;

import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.util.Locales;
import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngapi.NRanges;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.impl.BookSeriesAdv;
import org.zkoss.zss.ngmodel.impl.NameRefImpl;
import org.zkoss.zss.ngmodel.impl.RefImpl;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEvaluationContext;

/**
 * @author Pao
 */
public class FormulaEvalTest {

	@Before
	public void beforeTest() {
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
		DependencyTable table = ((BookSeriesAdv)book.getBookSeries()).getDependencyTable();

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

	private Ref toAreaRef(String sheet1, String sheet2, String area) {
		CellRegion r = new CellRegion(area);
		return new RefImpl("book1", sheet1, sheet2, r.getRow(), r.getColumn(), r.getLastRow(),
				r.getLastColumn());
	}

	private Ref toSheetRef(String sheet1) {
		return new RefImpl("book1", sheet1);
	}

	private Ref toCellRef(String sheet1, String sheet2, String cell) {
		CellRegion r = new CellRegion(cell);
		return new RefImpl("book1", sheet1, sheet2, r.getRow(), r.getColumn());
	}

	private Ref toNameRef(String name) {
		return new NameRefImpl("book1", null, name);
	}
	
	
	@Test
	public void testEvalAndModifyNormal(){
		NBook book = NBooks.createBook("book1");
		
		DependencyTable table = ((BookSeriesAdv)book.getBookSeries()).getDependencyTable();
		
		NSheet sheet1 = book.createSheet("Sheet1");
		NSheet sheet2 = book.createSheet("Sheet2");

		NCell cell = sheet1.getCell(0, 0);
		cell.setFormulaValue("Sheet2!A1");
		
		Assert.assertEquals(CellType.FORMULA, cell.getType());
		Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
		Assert.assertEquals(0D, cell.getValue());
		
		sheet2.getCell(0, 0).setValue("ABC");
		Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
		Assert.assertEquals(0D, cell.getValue());
		cell.clearFormulaResultCache();
		Assert.assertEquals(CellType.STRING, cell.getFormulaResultType());
		Assert.assertEquals("ABC", cell.getValue());
		
		sheet2.getCell(0, 0).setValue(Boolean.TRUE);
		Assert.assertEquals(CellType.STRING, cell.getFormulaResultType());
		Assert.assertEquals("ABC", cell.getValue());
		cell.clearFormulaResultCache();
		Assert.assertEquals(CellType.BOOLEAN, cell.getFormulaResultType());
		Assert.assertEquals(true, cell.getValue());
		
		sheet2.getCell(0, 0).setValue(123);
		Assert.assertEquals(CellType.BOOLEAN, cell.getFormulaResultType());
		Assert.assertEquals(true, cell.getValue());
		cell.clearFormulaResultCache();
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
		
		DependencyTable table = ((BookSeriesAdv)book.getBookSeries()).getDependencyTable();
		
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
		Assert.assertEquals(CellType.ERROR, cell.getFormulaResultType());
		Assert.assertEquals("#REF!", cell.getErrorValue().getErrorString());
		cell.clearFormulaResultCache();
		Assert.assertEquals(CellType.FORMULA, cell.getType());
		Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
		Assert.assertEquals(0D, cell.getValue());
		

		sheet2.getCell(0, 0).setValue("ABC");
		Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
		Assert.assertEquals(0D, cell.getValue());
		cell.clearFormulaResultCache();
		Assert.assertEquals(CellType.STRING, cell.getFormulaResultType());
		Assert.assertEquals("ABC", cell.getValue());
	}

	@Test
	public void testClearFormulaDependency() {
		NBook book = NBooks.createBook("book1");
		NSheet sheet1 = book.createSheet("Sheet1");
		DependencyTable table = ((BookSeriesAdv)book.getBookSeries()).getDependencyTable();

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

}
