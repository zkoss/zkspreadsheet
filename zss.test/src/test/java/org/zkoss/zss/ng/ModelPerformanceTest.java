/* EvaluationPerformanceTest.java

	Purpose:
		
	Description:
		
	History:
		Dec 19, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.ng;

import java.io.Closeable;
import java.io.InputStream;
import java.net.URL;
import java.util.Locale;
import java.util.Set;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.CellValue;
import org.zkoss.poi.xssf.usermodel.XSSFCell;
import org.zkoss.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.zkoss.poi.xssf.usermodel.XSSFRow;
import org.zkoss.poi.xssf.usermodel.XSSFSheet;
import org.zkoss.poi.xssf.usermodel.XSSFWorkbook;
import org.zkoss.util.Locales;
//import org.zkoss.zss.engine.RefSheet;
//import org.zkoss.zss.model.sys.XBook;
//import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.model.sys.impl.BookHelper;
//import org.zkoss.zss.model.sys.impl.XSSFBookImpl;
import org.zkoss.zss.ngapi.impl.imexp.ExcelImportFactory;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.impl.AbstractBookSeriesAdv;
import org.zkoss.zss.ngmodel.impl.RefImpl;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;

/**
 * @author Pao
 */
@Ignore
public class ModelPerformanceTest {
	private final static int ROW_COUNT = 1000;
	private final static int COL_COUNT = 50;

	private final static double EPSILON = 0.0000001;

	private final static URL file = ModelPerformanceTest.class.getResource("book/formula-1000x50.xlsx");
	private final static int NG_MODEL = 0;
	private final static int XBOOK_MODEL = 1;
	private final static int WORKBOOK_MODEL = 2;
	
	private XSSFFormulaEvaluator evaluator;
	
	@Before
	public void before() {
		memoryInit = -1;
		Locales.setThreadLocal(Locale.TAIWAN);
	}

	@Test
	public void testNGModel() {
		System.out.println("=== NG Model ===");
		testPerformanceAndMemory(NG_MODEL);
	}

	@Test
	public void testWorkbookModel() {
		System.out.println("=== WORKBOOK Model===");
		testPerformanceAndMemory(WORKBOOK_MODEL);
	}

	@Test
	public void testXBookModel() {
		System.out.println("=== XBook Model===");
		testPerformanceAndMemory(XBOOK_MODEL);
	}

	private void testPerformanceAndMemory(int type) {
		// base line
		System.out.println(">>> initial");
		showMemoryUsage();

		// create phase
		System.out.println(">>> creation");
		long start = System.currentTimeMillis();
		Object model = createModel(type);
		showSpendTime(start);
		showMemoryUsage();

		// evaluation
		System.out.println(">>> evaluation");
		start = System.currentTimeMillis();
		evaluation(type == NG_MODEL, model, 1.0);
		showSpendTime(start);
		showMemoryUsage();

		// evaluation with dependency tracking
		System.out.println(">>> modify and evaluation");
		start = System.currentTimeMillis();
		modifyFirstColumn(type == NG_MODEL, model, 2.0);
		evaluation(type == NG_MODEL, model, 2.0);
		showSpendTime(start);
		showMemoryUsage();
		
		// search dependencies
		System.out.println(">>> search dependencies");
		start = System.currentTimeMillis();
		searchDependencies(type, model);
		showSpendTime(start);
		showMemoryUsage();
	}

	private Object createModel(int type) {
		String bookName = "Book1";
		InputStream is = null;
		try {
			is = file.openStream();
			if(type == NG_MODEL) {
				NBook book = new ExcelImportFactory().createImporter().imports(is, bookName);
				book.getBookSeries().setAutoFormulaCacheClean(false);//in performance, we don't allow to clear automatically
				return book;
			} else if(type == WORKBOOK_MODEL) {
//				return new XSSFWorkbook(is);
			} else if(type == XBOOK_MODEL) {
//				return new XSSFBookImpl(bookName, is);
			}
		} catch(Exception e) {
			e.printStackTrace();
		} finally {
			close(is);
		}
		Assert.fail();
		return null;
	}

	private void close(Closeable r) {
		try {
			r.close();
		} catch(Exception e) {
		}
	}

	private void evaluation(boolean ngmodel, Object model, double firstColumnValue) {
		if(ngmodel) {
			NBook book = (NBook)model;
			NSheet sheet = book.getSheet(0);
			// get all values except first column
			for(int r = 0; r < ROW_COUNT; ++r) {
				double expected = firstColumnValue;
				for(int c = 1; c < COL_COUNT; ++c) {
					NCell cell = sheet.getCell(r, c);
					cell.clearFormulaResultCache();
					double v = cell.getNumberValue();
					Assert.assertEquals(expected, v, 0.0000001);
					expected *= 2.0;
				}
			}
		} else {
//			XSSFWorkbook book = (XSSFWorkbook)model;
//			if(book instanceof XBook){
//				evaluator = (XSSFFormulaEvaluator)((XBook)book).getFormulaEvaluator();
//				evaluator.clearAllCachedResultValues();
//			}else{
//				evaluator = XSSFFormulaEvaluator.create(book, null, null);
//			}
//			XSSFSheet sheet = book.getSheetAt(0);
//			// get all values except first column
//			for(int r = 0; r < ROW_COUNT; ++r) {
//				XSSFRow row = sheet.getRow(r);
//				double expected = firstColumnValue;
//				for(int c = 1; c < COL_COUNT; ++c) {
//					XSSFCell cell = row.getCell(c);
//					CellValue value = evaluator.evaluate(cell);
//					double v = value.getNumberValue();
//					Assert.assertEquals(expected, v, EPSILON);
//					expected *= 2.0;
//				}
//			}
		}
	}

	private void modifyFirstColumn(boolean ngmodel, Object model, double firstColumnValue) {
		if(ngmodel) {
			NBook book = (NBook)model;
			NSheet sheet = book.getSheet(0);
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
	
	private void searchDependencies(int type, Object model) {
		if(type == NG_MODEL) {
			NBook book = (NBook)model;
			AbstractBookSeriesAdv series = (AbstractBookSeriesAdv)book.getBookSeries();
			DependencyTable table = series.getDependencyTable();
			String bookName = book.getBookName();
			String sheetName = book.getSheet(0).getSheetName();
			for(int c = 0; c < COL_COUNT; ++c) {
				Ref ref = new RefImpl(bookName, sheetName, 0, c);
				Set<Ref> dependencies = table.getDependents(ref);
				Assert.assertEquals(COL_COUNT - c - 1, dependencies.size());
			}
		} else if(type == XBOOK_MODEL) {
//			XBook book = (XBook)model;
//			XSheet sheet = book.getWorksheetAt(0);
//			RefSheet refSheet = BookHelper.getRefSheet(book, sheet);
//			for(int c = 0; c < COL_COUNT; ++c) {
//				Set<org.zkoss.zss.engine.Ref>[] dependents = refSheet.getBothDependents(0, c);
//				Assert.assertEquals(COL_COUNT - c - 1, dependents[1].size());
//			}
		} else if(type == WORKBOOK_MODEL) {
			// do nothing, it has no dependency tracking
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
}
