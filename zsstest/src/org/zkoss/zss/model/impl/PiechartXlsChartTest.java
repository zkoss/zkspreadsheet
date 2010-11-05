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
import org.zkoss.poi.hssf.usermodel.HSSFChart;
import org.zkoss.poi.hssf.usermodel.HSSFSheet;
import org.zkoss.poi.hssf.usermodel.HSSFWorkbook;
import org.zkoss.poi.hssf.usermodel.HSSFChart.HSSFSeries;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Chart;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.model.Book;
import org.zkoss.zssex.model.impl.DrawingManager;

/**
 * Test chart anchor and type. 
 * @author henrichen
 *
 */
public class PiechartXlsChartTest {
	private Book _book;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename = "piechart.xls";
		final InputStream is = new ClassLocator().getResourceAsStream(filename);
		_book = (Book) new ExcelImporter().imports(is, filename);
		assertTrue(_book instanceof Book);
		assertTrue(_book instanceof HSSFBookImpl);
		assertTrue(_book instanceof HSSFWorkbook);
		assertEquals(filename, ((Book)_book).getBookName());
		assertEquals("Sheet1", _book.getSheetName(0));
		assertEquals(0, _book.getSheetIndex("Sheet1"));
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
		List<Chart> chartInfos = new DrawingManager(sheet1).getCharts();
		HSSFChart[] charts = HSSFChart.getSheetCharts(sheet1);
		assertEquals(2, charts.length);
		HSSFChart chart = charts[0];
		System.out.println("getChartHeight():"+chart.getChartHeight());
		System.out.println("getChartTitle():"+chart.getChartTitle());
		System.out.println("getgetChartWidth():"+chart.getChartWidth());
		System.out.println("getChartX():"+chart.getChartX());
		System.out.println("getChartY():"+chart.getChartY());
		HSSFSeries[] series = chart.getSeries();
		assertEquals(3, series.length);
		for(int j = 0; j < series.length; ++j) {
			HSSFSeries ser = series[j];
			System.out.println("--------"+j);
			System.out.println("ser.getNumValues():"+ser.getNumValues());
			System.out.println("ser.getSeriesTitle():"+ser.getSeriesTitle());
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
