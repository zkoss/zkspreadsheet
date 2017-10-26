package org.zkoss.zss.api.impl;

import static org.junit.Assert.*;

import java.util.Locale;

import org.junit.*;
import org.zkoss.zss.*;
import org.zkoss.zss.api.*;
import org.zkoss.zss.api.Range.AutoFillType;
import org.zkoss.zss.api.model.*;
import org.zkoss.zss.ui.impl.undo.*;
import org.zkoss.zss.ui.sys.UndoableAction;

/**
 * test autofill cells with different locked status in a protected sheet .
 * {@see Range} doesn't consider sheet protection, so can't test with Range API. 
 * @author hawk
 */
public class AutoFillActionTest{

	private static final String UNLOCKED = "unlocked";
	private Book book;
	private Sheet sheet;
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssLocale(Locale.TAIWAN);
		book = Util.loadBook(this, "book/unlocked.xlsx");
		sheet = book.getSheet(UNLOCKED);
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssLocale();
	}
		
	/**
	 * In a protected sheet,
	 * - auto-fill can fill among unlocked cells
	 */
	@Test
	public void fillUnlockedCells() {
		assertEquals(true, sheet.isProtected());
		UndoableAction action  = new AutoFillCellAction("autofill", sheet, 0, 1, 0, 2, 
				sheet, 0, 1, 0, 5, AutoFillType.DEFAULT);
		action.doAction();
		
		for (int column = 3; column < 6 ; column++){
			Range targetCell = Ranges.range(sheet, 0, column);
			assertEquals(false, targetCell.getCellStyle().isLocked());
			assertEquals(Integer.toString(column+1), targetCell.getCellEditText());
		}
	}

	/**
	 * if auto-fill destination region contains at least 1 locked cell, it doesn't fill any cell
	 */
	@Test
	public void fill2OneLockedCell() {
		assertEquals(true, sheet.isProtected());
		UndoableAction action  = new AutoFillCellAction("autofill", sheet, 0, 1, 0, 2, 
				sheet, 0, 1, 0, 6, AutoFillType.DEFAULT);
		action.doAction();
		int column;
		for (column = 3; column < 6  ; column++){
			Range targetCell = Ranges.range(sheet, 0, column);
			assertEquals(false, targetCell.getCellStyle().isLocked());
			assertEquals("", targetCell.getCellEditText());
		}
		Range targetCell = Ranges.range(sheet, 0, column);
		assertEquals(true, targetCell.getCellStyle().isLocked());
		assertEquals("", targetCell.getCellEditText());
	}
	

	/**
	 * a destination range excludes the selected source cells.
	 * Select A1:B1 and drag to C1:D1, the destination is C1:D1.
	 * so users can fill from locked cells to unlocked cells
	 */
	@Ignore("not implemented yet")
	public void dragFromLockedCells() {
		assertEquals(true, sheet.isProtected());
		
	}
	
}
