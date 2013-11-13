package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;

import java.util.Locale;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

/**
 * @author Hawk
 *
 */
public class Issue500Test {
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssContextLocale(Locale.TAIWAN);
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssContextLocale();
	}
	
	
	@Test
	public void testZSS502_2003(){
		Book book = Util.loadBook(this,"book/502-crossSheetReference.xls");
		Sheet sheet = book.getSheet("cell-reference");
		Range referencingCell = Ranges.range(sheet, "C4");
		assertEquals("=row!A1", referencingCell.getCellEditText());
		assertEquals("The first row is freezed.", referencingCell.getCellFormatText());

		Ranges.range(book.getSheet("row")).deleteSheet();
		
		assertEquals("='#REF'!A1", referencingCell.getCellEditText());
		//because ZSS-474 is unresolved, the assert below won't pass 
		//assertEquals("#REF", referencingCell.getCellFormatText());
	}
	
	@Test
	public void testZSS502_2007(){
		Book book = Util.loadBook(this,"book/502-crossSheetReference.xlsx");
		Sheet sheet = book.getSheet("cell-reference");
		Range referencingCell = Ranges.range(sheet, "C4");
		assertEquals("=row!A1", referencingCell.getCellEditText());
		assertEquals("The first row is freezed.", referencingCell.getCellFormatText());

		Ranges.range(book.getSheet("row")).deleteSheet();
		
		assertEquals("=row!A1", referencingCell.getCellEditText());
		//because ZSS-474 is unresolved, the assert below won't pass 
//		assertEquals("#REF", referencingCell.getCellFormatText());
	}
	
	
	@Test
	public void testZSS502_NonExistingSheet2003(){
		Book book = Util.loadBook(this,"book/blank.xls");
		Sheet sheet = book.getSheetAt(0);
		Range cell = Ranges.range(sheet, "A1");
		cell.setCellEditText("=nonExisted!B1");
		assertEquals("='#REF'!B1", cell.getCellEditText());
		assertEquals("#REF!", cell.getCellFormatText());
	}
	
	@Test
	public void testZSS502_NonExistingSheet2007(){
		Book book = Util.loadBook(this,"book/blank.xlsx");
		Sheet sheet = book.getSheetAt(0);
		Range cell = Ranges.range(sheet, "A1");
		cell.setCellEditText("=nonExisted!B1");
		assertEquals("=nonExisted!B1", cell.getCellEditText());
		assertEquals("#REF!", cell.getCellFormatText());
	}
	
	@Test
	public void testZSS510() {
		Book book = Util.loadBook(this, "book/blank.xlsx");
		Sheet sheet = book.getSheetAt(0);
		Range r = Ranges.range(sheet, "A1");
		r.setCellEditText("Hello");
		CellOperationUtil.applyDataFormat(r, "");
		r.getCellFormatText(); // get text shouldn't cause IndexOutBoundaryException
		assertEquals("General", r.getCellStyle().getDataFormat()); // should get General instead of empty string
	}
}
