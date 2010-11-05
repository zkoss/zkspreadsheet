/* FormulaEvaluatorTest.java

	Purpose:
		
	Description:
		
	History:
		Mar 17, 2010 11:55:42 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.impl;


import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.InputStream;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.xssf.usermodel.XSSFSheet;
import org.zkoss.poi.xssf.usermodel.XSSFWorkbook;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.model.Book;
import org.zkoss.zssex.model.impl.DrawingManager;

/**
 * Test chart anchor and type. 
 * @author henrichen
 *
 */
public class ImageXlsxPictureTest {
	private Book _book;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename = "image.xlsx";
		final InputStream is = new ClassLocator().getResourceAsStream(filename);
		_book = (Book) new ExcelImporter().imports(is, filename);
		assertTrue(_book instanceof Book);
		assertTrue(_book instanceof XSSFBookImpl);
		assertTrue(_book instanceof XSSFWorkbook);
		assertEquals(filename, ((Book)_book).getBookName());
		assertEquals("Sheet1", _book.getSheetName(0));
		assertEquals("Sheet2", _book.getSheetName(1));
		assertEquals("Sheet3", _book.getSheetName(2));
		assertEquals(0, _book.getSheetIndex("Sheet1"));
		assertEquals(1, _book.getSheetIndex("Sheet2"));
		assertEquals(2, _book.getSheetIndex("Sheet3"));
	}

	/**
	 * @throws java.lang.Exception
	 */
	@After
	public void tearDown() throws Exception {
		_book = null;
	}

	@Test
	public void testPictures() {
		XSSFSheet sheet1 = (XSSFSheet)_book.getSheetAt(0);
		new DrawingManager(sheet1).getPictures();
		_book.getAllPictures();
	}
	
	private void testToFormulaString(Cell cell, String expect) {
/*		EvaluationCell srcCell = HSSFEvaluationTestHelper.wrapCell((HSSFCell)cell);
		HSSFEvaluationWorkbook evalbook = HSSFEvaluationWorkbook.create((HSSFWorkbook)_book);
		Ptg[] ptgs = evalbook.getFormulaTokens(srcCell);
		final String formula = FormulaRenderer.toFormulaString(evalbook, ptgs);
		assertEquals(expect, formula);
*/	}
}
