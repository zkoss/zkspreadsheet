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
import java.util.List;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.lang.Library;
import org.zkoss.poi.hssf.usermodel.HSSFChart;
import org.zkoss.poi.hssf.usermodel.HSSFSheet;
import org.zkoss.poi.hssf.usermodel.HSSFWorkbook;
import org.zkoss.poi.hssf.usermodel.HSSFChart.HSSFSeries;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Chart;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.model.Book;
import org.zkoss.zssex.model.impl.DrawingManagerImpl;

/**
 * Test chart anchor and type. 
 * @author henrichen
 *
 */
public class ColumnchartXlsChartTest {
	private Book _book;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename = "columnchart.xls";
		final InputStream is = new ClassLocator().getResourceAsStream(filename);
		_book = (Book) new ExcelImporter().imports(is, filename);
		assertTrue(_book instanceof Book);
		assertTrue(_book instanceof HSSFBookImpl);
		assertTrue(_book instanceof HSSFWorkbook);
		assertEquals(filename, ((Book)_book).getBookName());
		assertEquals("Sheet1", _book.getSheetName(0));
		assertEquals(0, _book.getSheetIndex("Sheet1"));
		Library.setProperty("org.zkoss.zss.model.EscherAggregate.class", "org.zkoss.zssex.model.impl.ZKEscherAggregate");
		Library.setProperty("org.zkoss.zss.model.EscherAggregate.UTEST.class", "org.zkoss.zssex.model.impl.ZKEscherAggregate");
	}

	/**
	 * @throws java.lang.Exception
	 */
	@After
	public void tearDown() throws Exception {
		_book = null;
	}

	@Test
	public void testColumnchart() {
		HSSFSheet sheet1 = (HSSFSheet) _book.getSheet("Sheet1");
		List<Chart> chartXes = new DrawingManagerImpl(sheet1).getCharts();
		assertEquals(1, chartXes.size());
		HSSFChart chart = (HSSFChart) chartXes.get(0).getChartInfo();
		assertEquals("2003-2006 Income Summary", chart.getChartTitle());
		HSSFSeries[] series = chart.getSeries();
		assertEquals(3, series.length);
		
		HSSFSeries ser = series[0];
		assertEquals(4, ser.getNumValues());
		assertEquals("Total Revenues", ser.getSeriesTitle());

		ser = series[1];
		assertEquals(4, ser.getNumValues());
		assertEquals("Total Expenses", ser.getSeriesTitle());
		
		ser = series[2];
		assertEquals(4, ser.getNumValues());
		assertEquals("Profit/Loss", ser.getSeriesTitle());
		
		for(int j = 0; j < series.length; ++j) {
			System.out.println("--------"+j);
			System.out.println("ser.getValueType():"+ser.getValueType());
			System.out.println("ser.getDataCategoryLabels():"+ser.getDataCategoryLabels());
			System.out.println("ser.getDataName():"+ser.getDataName());
			System.out.println("ser.getDataSecondaryCategoryLabels():"+ser.getDataSecondaryCategoryLabels());
			System.out.println("ser.getDataValues():"+ser.getDataValues());
			System.out.println("ser.getSeries():"+ser.getSeries());
		}
	}
	
	private void testToFormulaString(Cell cell, String expect) {
/*		EvaluationCell srcCell = HSSFEvaluationTestHelper.wrapCell((HSSFCell)cell);
		HSSFEvaluationWorkbook evalbook = HSSFEvaluationWorkbook.create((HSSFWorkbook)_book);
		Ptg[] ptgs = evalbook.getFormulaTokens(srcCell);
		final String formula = FormulaRenderer.toFormulaString(evalbook, ptgs);
		assertEquals(expect, formula);
*/	}
}
