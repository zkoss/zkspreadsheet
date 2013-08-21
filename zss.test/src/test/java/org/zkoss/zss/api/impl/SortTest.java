package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;

import java.io.InputStream;
import java.util.Arrays;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

public class SortTest {
	
	private static Book _workbook;
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void setUp() throws Exception {
		final String filename = "book/blank.xlsx";
		final InputStream is = PasteTest.class.getResourceAsStream(filename);
		_workbook = Importers.getImporter().imports(is, filename);
	}
	
	@After
	public void tearDown() throws Exception {
		_workbook = null;
	}
	
	/**
	 * sort random number from 1:100
	 */
	@Test
	public void simpleSort() {
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		int[] rands = new int[100];
		for(int i = 0; i < 100; i++) {
			rands[i] = (int)Math.random() * 100 + 1;
			Ranges.range(sheet1, i, 0).setCellValue(rands[i]);
		}
		Ranges.range(sheet1, 0, 0, 99, 0).sort(false);
		Arrays.sort(rands);
		for(int i = 0; i < 100; i++) {
			assertEquals(rands[i], Ranges.range(sheet1, i, 0).getCellData().getDoubleValue(), 1E-8);
		}
	}

}
