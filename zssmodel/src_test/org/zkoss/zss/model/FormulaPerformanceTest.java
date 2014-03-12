/* EvaluationPerformanceTest.java

	Purpose:
		
	Description:
		
	History:
		Dec 19, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.model;

import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.CellValue;
import org.zkoss.poi.ss.util.AreaReference;
import org.zkoss.poi.ss.util.CellReference;
import org.zkoss.poi.xssf.usermodel.XSSFCell;
import org.zkoss.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.zkoss.poi.xssf.usermodel.XSSFRow;
import org.zkoss.poi.xssf.usermodel.XSSFSheet;
import org.zkoss.poi.xssf.usermodel.XSSFWorkbook;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.impl.FormulaCacheCleaner;

/**
 * @author Pao
 */
@Ignore
public class FormulaPerformanceTest {
	private final static int ROW_COUNT = 1000;
	private final static int COL_COUNT = 50;
	private static final double EPSILON = 0.0000001;

	private XSSFFormulaEvaluator evaluator;
	@BeforeClass
	static public void beforeClass() {
		Setup.touch();
	}
	@Before
	public void before() {
		memoryInit = -1;
	}

	@Test
	public void testPOIModel() {
		System.out.println("=== ZPOI Model===");
		testPerformanceAndMemory(false);
	}

	
	
	@Test
	public void testNGModel() {
		System.out.println("=== NG Model ===");
		testPerformanceAndMemory(true);
	}
	
//	@Test
	public void testNGMemory(){
		boolean ngmodel = true;
		Object book = createModel(ngmodel);
		evaluation(ngmodel, book, 1.0);
		modifyFirstColumn(ngmodel, book, 2.0);
		evaluation(ngmodel, book, 2.0);
		//
		System.out.println(">>>>>>> go and use MAT");
		try {
			Thread.sleep(200000);
		} catch (InterruptedException e) {e.printStackTrace();
		}
	}

	private void testPerformanceAndMemory(boolean ngmodel) {
		// base line
		System.out.println(">>> initial");
		showMemoryUsage();

		// create phase
		System.out.println(">>> creation");
		long start = System.currentTimeMillis();
		Object book = createModel(ngmodel);
		showSpendTime(start);
		showMemoryUsage();

		// evaluation
		System.out.println(">>> evaluation");
		start = System.currentTimeMillis();
		evaluation(ngmodel, book, 1.0);
		showSpendTime(start);
		showMemoryUsage();

		// evaluation with dependency tracking
		System.out.println(">>> modify and evaluation");
		start = System.currentTimeMillis();
		modifyFirstColumn(ngmodel, book, 2.0);
		evaluation(ngmodel, book, 2.0);
		showSpendTime(start);
		showMemoryUsage();
	}

	private Object createModel(boolean ngmodel) {
		if(ngmodel) {
			SBook book = SBooks.createBook("Book1");
			book.getBookSeries().setAutoFormulaCacheClean(false);//in performance, we don't allow to clear automatically
			SSheet sheet = book.createSheet("Sheet1");
			// first column
			modifyFirstColumn(ngmodel, book, 1.0);
			// other columns
			for(int r = 0; r < ROW_COUNT; ++r) {
				for(int c = 1; c < COL_COUNT; ++c) {
					String region = new AreaReference(new CellReference(r, 0), new CellReference(r, c - 1))
							.formatAsString();
					sheet.getCell(r, c).setFormulaValue("SUM(" + region + ")");
				}
			}
			return book;
		} else {
			XSSFWorkbook book = new XSSFWorkbook();
			XSSFSheet sheet = book.createSheet("Sheet1");
			evaluator = XSSFFormulaEvaluator.create(book, null, null);
			// first column
			modifyFirstColumn(ngmodel, book, 1.0);
			// other columns
			for(int r = 0; r < ROW_COUNT; ++r) {
				XSSFRow row = sheet.getRow(r);
				if(row == null) {
					row = sheet.createRow(r);
				}
				for(int c = 1; c < COL_COUNT; ++c) {
					String region = new AreaReference(new CellReference(r, 0), new CellReference(r, c - 1))
							.formatAsString();
					XSSFCell cell = row.getCell(c);
					if(cell == null) {
						cell = row.createCell(c);
					}
					cell.setCellFormula("SUM(" + region + ")");
				}
			}
			return book;
		}
	}

	private void evaluation(boolean ngmodel, Object model, double firstColumnValue) {
		if(ngmodel) {
			SBook book = (SBook)model;
			SSheet sheet = book.getSheet(0);
			// get all values except first column
			for(int r = 0; r < ROW_COUNT; ++r) {
				double expected = firstColumnValue;
				for(int c = 1; c < COL_COUNT; ++c) {
					SCell cell = sheet.getCell(r, c);
					cell.clearFormulaResultCache(); // cell clear it's cache automatically
					double v = cell.getNumberValue();
					Assert.assertEquals(expected, v, 0.0000001);
					expected *= 2.0;
				}
			}
		} else {
			XSSFWorkbook book = (XSSFWorkbook)model;
			evaluator = XSSFFormulaEvaluator.create(book, null, null);
			XSSFSheet sheet = book.getSheetAt(0);
			// get all values except first column
			for(int r = 0; r < ROW_COUNT; ++r) {
				XSSFRow row = sheet.getRow(r);
				double expected = firstColumnValue;
				for(int c = 1; c < COL_COUNT; ++c) {
					XSSFCell cell = row.getCell(c);
					CellValue value = evaluator.evaluate(cell);
					double v = value.getNumberValue();
					Assert.assertEquals(expected, v, EPSILON);
					expected *= 2.0;
				}
			}
		}
	}

	private void modifyFirstColumn(boolean ngmodel, Object model, double firstColumnValue) {
		if(ngmodel) {
			SBook book = (SBook)model;
			SSheet sheet = book.getSheet(0);
			for(int r = 0; r < ROW_COUNT; ++r) {
				sheet.getCell(r, 0).setNumberValue(firstColumnValue);
			}
		} else {
			XSSFWorkbook book = (XSSFWorkbook)model;
			XSSFSheet sheet = book.getSheetAt(0);
			for(int r = 0; r < ROW_COUNT; ++r) {
				XSSFRow row = sheet.getRow(r);
				if(row == null) {
					row = sheet.createRow(r);
				}
				XSSFCell cell = row.getCell(0);
				if(cell == null) {
					cell = row.createCell(0);
				}
				cell.setCellValue(firstColumnValue);
			}
		}
	}

	private void showSpendTime(long lastTimeInNS) {
		long now = System.currentTimeMillis();
		long spend = now - lastTimeInNS;
		System.out.printf("spend: %,16d ms\n", spend);
	}

	long memoryInit = -1;

	private void showMemoryUsage() {
		long ms = System.currentTimeMillis();
		Runtime rt = Runtime.getRuntime();
		long last = rt.freeMemory();
		while(true) {
			rt.gc();
			Thread.yield();
			if(rt.freeMemory() == last) {
				break;
			}
			last = rt.freeMemory();
		}
		ms = System.currentTimeMillis() - ms;
		long usedHeap;
		if(memoryInit == -1) {
			memoryInit = usedHeap = rt.totalMemory() - rt.freeMemory();
		} else {
			usedHeap = rt.totalMemory() - rt.freeMemory() - memoryInit;
		}
		System.out.printf("used:  %,16d bytes (GC spend %,ds)\n", usedHeap, ms / 1000L);
	}

	@Test
	public void testDepend() {
		SBook book = SBooks.createBook("Book1");
		SSheet sheet = book.createSheet("Sheet1");
		sheet.getCell(0, 0).setNumberValue(3.0); // A1
		sheet.getCell(0, 1).setFormulaValue("A1 * 2"); // B1
		Assert.assertEquals(3.0, sheet.getCell(0, 0).getNumberValue(), EPSILON);
		Assert.assertEquals(6.0, sheet.getCell(0, 1).getNumberValue(), EPSILON);

		sheet.getCell(0, 0).setNumberValue(1.0);
		Assert.assertEquals(6.0, sheet.getCell(0, 1).getNumberValue(), EPSILON);
//		sheet.getCell(0, 1).clearFormulaResultCache();// cell clear it's cache automatically
		Assert.assertEquals(2.0, sheet.getCell(0, 1).getNumberValue(), EPSILON);

		sheet.getCell(0, 0).clearValue();
//		sheet.getCell(0, 1).clearFormulaResultCache();// cell clear it's cache automatically
		Assert.assertEquals(0.0, sheet.getCell(0, 1).getNumberValue(), EPSILON);
	}
}
