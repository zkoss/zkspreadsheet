package org.zkoss.zss.api.impl;

import static org.junit.Assert.*;

import java.util.Locale;

import org.junit.*;
import org.zkoss.zss.*;
import org.zkoss.zss.api.*;
import org.zkoss.zss.api.model.*;
import org.zkoss.zss.ui.impl.undo.*;
import org.zkoss.zss.ui.impl.undo.ClearCellAction.Type;
import org.zkoss.zss.ui.sys.UndoableAction;

/**
 * test copy/cut cells with different locked status in a protected sheet .
 * {@see Range} doesn't consider sheet protection, so can't test with Range API. 
 * @author hawk
 */
public class CopyCutActionTest{

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
	 * - users can clearFormat among unlocked cells
	 */
	@Test	
	public void clearFormatUnlockedCells() {
		assertEquals(true, sheet.isProtected());
		UndoableAction action  = new ClearCellAction("clear", sheet, 0, 0, 1, 2, ClearCellAction.Type.STYLE);
		action.doAction();
		
		for (int row = 0; row < 2; row++) {
			for (int column = 0 ; column <3 ; column++){
				final int v = row * 3 + column + 1; 
				Range srcCell = Ranges.range(sheet, row, column);
				assertEquals(true, srcCell.getCellStyle().isLocked());
				assertEquals(v, ((Number)srcCell.getCellValue()).intValue());
			}
		}
		
	}
	/**
	 * In a protected sheet,
	 * - users can NOT clearFormat among locked cells
	 */
	@Test	
	public void clearFormatLockedCells() {
		assertEquals(true, sheet.isProtected());
		UndoableAction action  = new ClearCellAction("clear", sheet, 0, 0, 2, 2, ClearCellAction.Type.STYLE);
		action.doAction();
		for (int row = 0; row < 2; row++) {
			for (int column = 0 ; column <3 ; column++){
				final int v = row * 3 + column + 1; 
				Range srcCell = Ranges.range(sheet, row, column);
				assertEquals(false, srcCell.getCellStyle().isLocked());
				assertEquals(v, ((Number)srcCell.getCellValue()).intValue());
			}
		}
		
	}
	/**
	 * In a protected sheet,
	 * - users can clearContent among unlocked cells
	 */
	@Test	
	public void clearContentUnlockedCells() {
		assertEquals(true, sheet.isProtected());
		UndoableAction action  = new ClearCellAction("clear", sheet, 0, 0, 1, 2, ClearCellAction.Type.CONTENT);
		action.doAction();
		
		for (int row = 0; row < 2; row++) {
			for (int column = 0 ; column <3 ; column++){
				Range srcCell = Ranges.range(sheet, row, column);
				assertEquals(false, srcCell.getCellStyle().isLocked());
				assertEquals(null, srcCell.getCellValue());
			}
		}
	}
	/**
	 * In a protected sheet,
	 * - users can NOT clearContent among locked cells
	 */
	@Test	
	public void clearContentLockedCells() {
		assertEquals(true, sheet.isProtected());
		UndoableAction action  = new ClearCellAction("clear", sheet, 0, 0, 2, 2, ClearCellAction.Type.CONTENT);
		action.doAction();
		
		for (int row = 0; row < 2; row++) {
			for (int column = 0 ; column <3 ; column++){
				final int v = row * 3 + column + 1; 
				Range srcCell = Ranges.range(sheet, row, column);
				assertEquals(false, srcCell.getCellStyle().isLocked());
				assertEquals(v, ((Number)srcCell.getCellValue()).intValue());
			}
		}
	}
	/**
	 * In a protected sheet,
	 * - users can clearAll among unlocked cells
	 */
	@Test	
	public void clearAllUnlockedCells() {
		assertEquals(true, sheet.isProtected());
		UndoableAction action  = new ClearCellAction("clear", sheet, 0, 0, 1, 2, ClearCellAction.Type.ALL);
		action.doAction();
		
		for (int row = 0; row < 2; row++) {
			for (int column = 0 ; column <3 ; column++){
				Range srcCell = Ranges.range(sheet, row, column);
				assertEquals(true, srcCell.getCellStyle().isLocked());
				assertEquals(null, srcCell.getCellValue());
			}
		}
	}
	/**
	 * In a protected sheet,
	 * - users can NOT clearAll among locked cells
	 */
	@Test	
	public void clearAllLockedCells() {
		assertEquals(true, sheet.isProtected());
		UndoableAction action  = new ClearCellAction("clear", sheet, 0, 0, 2, 2, ClearCellAction.Type.ALL);
		action.doAction();
		
		for (int row = 0; row < 2; row++) {
			for (int column = 0 ; column <3 ; column++){
				final int v = row * 3 + column + 1; 
				Range srcCell = Ranges.range(sheet, row, column);
				assertEquals(false, srcCell.getCellStyle().isLocked());
				assertEquals(v, ((Number)srcCell.getCellValue()).intValue());
			}
		}
	}
	/**
	 * In a protected sheet,
	 * - users can copy among unlocked cells
	 */
	@Test
	public void copyUnlockedCells() {
		assertEquals(true, sheet.isProtected());
		UndoableAction action  = new PasteCellAction("paste", sheet, 0, 0, 0, 2, false, false,
				sheet, 1, 0, 1, 2, false, false);
		action.doAction();
		
		for (int column = 0 ; column <3 ; column++){
			Range srcCell = Ranges.range(sheet, 0, column);
			Range targetCell = Ranges.range(sheet, 1, column);
			assertEquals(false, srcCell.getCellStyle().isLocked());
			assertEquals(false, targetCell.getCellStyle().isLocked());
			assertEquals(srcCell.getCellEditText(), targetCell.getCellEditText());
		}
	}

	@Test
	public void copy2LockedCells() {
		assertEquals(true, sheet.isProtected());
		UndoableAction action  = new PasteCellAction("paste", sheet, 0, 0, 0, 2, false, false,
				sheet, 2, 0, 2, 2, false, false);
		action.doAction();
		
		for (int column = 0 ; column <3 ; column++){
			Range srcCell = Ranges.range(sheet, 0, column);
			Range targetCell = Ranges.range(sheet, 2, column);
			assertEquals(false, srcCell.getCellStyle().isLocked());
			assertEquals(true, targetCell.getCellStyle().isLocked());
			assertNotEquals(srcCell.getCellEditText(), targetCell.getCellEditText());
		}
	}

	/**
	 * In a protected sheet,
	 *  - users can copy a locked cell and paste it to an unlocked one.
	 * - it doesn't copy the "locked" status
	 */
	@Test
	public void copyLockedCells() {
		assertEquals(true, sheet.isProtected());
		
		UndoableAction action  = new PasteCellAction("paste", sheet, 2, 0, 2, 2, false, false,
				sheet, 0, 0, 0, 2, false, false);
		action.doAction();
		
		for (int column = 0 ; column <3 ; column++){
			Range srcCell = Ranges.range(book.getSheet(UNLOCKED), 2, column);
			Range targetCell = Ranges.range(book.getSheet(UNLOCKED), 0, column);
			assertEquals(true, srcCell.getCellStyle().isLocked());
			assertEquals(false, targetCell.getCellStyle().isLocked());
			assertEquals(srcCell.getCellEditText(), targetCell.getCellEditText());
		}
	}
	
	/**
	 * if there is a locked cell in the affected target range while copying, abort the pasting action
	 */
	@Test
	public void copy2OneLockedCell() {
		assertEquals(true, sheet.isProtected());
		UndoableAction action  = new PasteCellAction("paste", sheet, 0, 0, 0, 2, false, false,
				sheet, 1, 1, 1, 3, false, false);
		action.doAction();
		
		for (int column = 0 ; column <2 ; column++){
			Range srcCell = Ranges.range(sheet, 0, column);
			Range targetCell = Ranges.range(sheet, 1, column+1);
			assertEquals(false, srcCell.getCellStyle().isLocked());
			assertEquals(false, targetCell.getCellStyle().isLocked());
			assertNotEquals(srcCell.getCellEditText(), targetCell.getCellEditText());
		}
		Range srcCell = Ranges.range(sheet, 0, 2);
		Range targetCell = Ranges.range(sheet, 1, 3);
		assertEquals(false, srcCell.getCellStyle().isLocked());
		assertEquals(true, targetCell.getCellStyle().isLocked());
		assertNotEquals(srcCell.getCellEditText(), targetCell.getCellEditText());
	}
	
	/**
	 * cut an unlocked cell, the cell should keep unlocked.
	 * without sheet protection, cut an unlocked cell, the cell restores to the default value, locked is true
	 */
	@Test
	public void cutUnlockedCell() {
		assertEquals(true, sheet.isProtected());
		String[] cellText = new String[3];
		for (int column = 0 ; column <3 ; column++){
			Range srcCell = Ranges.range(sheet, 0, column);
			cellText[column] = srcCell.getCellEditText();
		}

		UndoableAction action  = new CutCellAction("cut", sheet, 0, 0, 0, 2, false, false,
				sheet, 1, 0, 1, 2, false, false);
		action.doAction();
		
		for (int column = 0 ; column <3 ; column++){
			Range srcCell = Ranges.range(sheet, 0, column);
			Range targetCell = Ranges.range(sheet, 1, column);
			assertEquals(false, srcCell.getCellStyle().isLocked());
			assertEquals(false, targetCell.getCellStyle().isLocked());
			assertEquals("", srcCell.getCellEditText());
			assertEquals(cellText[column], targetCell.getCellEditText());
		}
	}
}
