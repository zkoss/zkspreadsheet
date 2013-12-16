/* FormulaEvalTest.java

	Purpose:
		
	Description:
		
	History:
		Dec 9, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.ngmodel;

import java.util.Locale;

import junit.framework.Assert;

import org.junit.Before;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.util.Locales;
import org.zkoss.zss.ngmodel.impl.BookImpl;

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
		NBook book = new BookImpl("book1");
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
		NBook book = new BookImpl("book1");
		NSheet sheet1 = book.createSheet("Sheet1");
		NSheet sheet2 = book.createSheet("Sheet2");
		NSheet sheetA = book.createSheet("SheetA");

		// cells in sheet1
		// basic formula
		sheet1.getCell(0, 0).setFormulaValue("B1");
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

		// error test
		sheet1.getCell(12, 0).setFormulaValue("SUM(NotExistedName)");

		sheet1.getCell(0, 1).setStringValue("HELLO!!!");
		sheet1.getCell(1, 1).setStringValue("ZK!!!");
		for(int i = 0; i < 10; ++i) {
			sheet1.getCell(i, 2).setNumberValue((double)i);
			sheet1.getCell(i, 3).setNumberValue(i * 10.0);
			sheet1.getCell(i, 4).setNumberValue(i * 10.0);
		}

		// sheet2
		// cells in sheet2
		for(int i = 0; i < 10; ++i) {
			sheet2.getCell(i, 3).setNumberValue(i * 10.0);
			sheet2.getCell(i, 4).setNumberValue(i * 10.0);
		}

		// sheetA
		// cells in sheetA
		for(int i = 0; i < 10; ++i) {
			sheetA.getCell(i, 3).setNumberValue(i * 10.0);
			sheetA.getCell(i, 4).setNumberValue(i * 10.0);
		}
		
		// named range
		// book.createName("ABC", "Sheet1").setRefersToFormula("C1:C10"); // TODO ZSS 3.5
		book.createName("ABC").setRefersToFormula("Sheet1!C1:C10");
		book.createName("DEF").setRefersToFormula("SheetA!D1:D10");
		book.createName("GHI").setRefersToFormula("Sheet1:Sheet2!C1:C10");

		// test
		String[] expected = new String[]{"HELLO!!!", "ZK!!!", "45.0", "22.5", "67.5", "4.5", "0.0", "10.0",
				"45.0", "45.0", "450.0", (450 * 6) + ".0", "#NAME?"};
		for(int i = 0; i <= 99; ++i) {
			NCell cell = sheet1.getCell(i, 0);
			if(cell.isNull()) {
				break;
			}

//			System.out.println(cell.getType() == CellType.FORMULA ? cell.getFormulaValue() : cell.getValue()
//					.toString());

			switch(cell.getFormulaResultType()) {
				case STRING:
					Assert.assertEquals(expected[i], cell.getStringValue());
					break;
				case NUMBER:
					Assert.assertEquals(expected[i], cell.getNumberValue().toString());
					break;
				case ERROR:
					Assert.assertEquals(expected[i], cell.getErrorValue().gettErrorString());
					break;
				default:
					Assert.fail("not expected result type: " + cell.getFormulaResultType());
					break;
			}
		}
	}

}
