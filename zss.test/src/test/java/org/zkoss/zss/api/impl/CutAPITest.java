package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;

import java.io.InputStream;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellData.CellType;
import org.zkoss.zss.api.model.Sheet;

/**
 * test cut API
 * validation: 
 * 1. destination is correct
 * 2. source is cleared
 * @author kuro
 *
 */
public class CutAPITest {
	
	private static Book _workbook;
	// 3 x 3, 1-9 matrix (source)
	// Source Range (H11, J13)
	private static 	int SRC_TOP_ROW = 10;
	private static int SRC_LEFT_COL = 7;
	private static int SRC_BOTTOM_ROW = 12;
	private static int SRC_RIGHT_COL = 9;
	private static int SRC_ROW_COUNT = SRC_BOTTOM_ROW - SRC_TOP_ROW + 1;
	private static int SRC_COL_COUNT = SRC_RIGHT_COL - SRC_LEFT_COL + 1;
	
	@Before
	public void setUp() throws Exception {
		Setup.touch();
		final String filename = "overlappingPaste.xlsx";
		final InputStream is = new ClassLocator().getResourceAsStream(filename);
		_workbook = Importers.getImporter().imports(is, filename);
	}
	
	@After
	public void tearDown() throws Exception {
		_workbook = null;
	}
	
	/**
	 * source (H11,J13) to destination (C4, E6)
	 */
	@Test
	public void testCutPasteWithMerge() {
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		
		// preparation
		// Merge (H12, J12)
		Range range_H12J12 = Ranges.range(sheet1, 11, 7, 11, 9);
		range_H12J12.merge(false);
		
		int dstTopRow = 3;
		int dstLeftCol = 2;
		int dstBottomRow = 5;
		int dstRightCol = 4;
		
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		srcRange.paste(dstRange, true);
		
		// validate destination
//		assertEquals(1, Ranges.range(sheet1, row, col).getCellData().getDoubleValue(), 1E-8);
	}
	
	@Test
	public void testCutPasteWithMergeOverlap() {
		
	}
	
	/**
	 * source (H11,J13) to destination (G10, I12)
	 */
	@Test
	public void testCutPasteToLeftTop() {

		int dstTopRow = 9;
		int dstLeftCol = 6;
		int dstBottomRow = 11;
		int dstRightCol = 8;
		
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange, true);
		
		int realDstRowCount = pasteRange.getLastRow() - pasteRange.getRow();
		int realDstColCount = pasteRange.getLastColumn() - pasteRange.getColumn();
		
		// validate destination
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / SRC_ROW_COUNT, realDstColCount / SRC_COL_COUNT);
		
		// validate source (is clean?)
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 10, 9).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 11, 9).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 12, 7).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 12, 8).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 12, 9).getCellData().getType().ordinal(), 1E-8);
	}
	
	/**
	 * source (H11,J13) to destination (J13, L15)
	 */
	@Test
	public void testCutPasteToRightBottom() {

		int dstTopRow = 12;
		int dstLeftCol = 9;
		int dstBottomRow = 14;
		int dstRightCol = 11;
		
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange, true);
		
		int realDstRowCount = pasteRange.getLastRow() - pasteRange.getRow();
		int realDstColCount = pasteRange.getLastColumn() - pasteRange.getColumn();
		
		// validate destination
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / SRC_ROW_COUNT, realDstColCount / SRC_COL_COUNT);
		
		// validate source (is clean?)
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 10, 7).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 10, 8).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 10, 9).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 11, 7).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 11, 8).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 11, 9).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 12, 7).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 12, 8).getCellData().getType().ordinal(), 1E-8);
	}
	
	/**
	 * source (H11,J13) to destination (G9, I14)
	 */
	@Test
	public void testCutPasteToLeft() {
		
		int dstTopRow = 8;
		int dstLeftCol = 6;
		int dstBottomRow = 13;
		int dstRightCol = 8;
		
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange, true);
		
		int realDstRowCount = pasteRange.getLastRow() - pasteRange.getRow();
		int realDstColCount = pasteRange.getLastColumn() - pasteRange.getColumn();
		
		// validate destination
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / SRC_ROW_COUNT, realDstColCount / SRC_COL_COUNT);
		
		// validate source (is clean?)
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 10, 9).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 11, 9).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 12, 9).getCellData().getType().ordinal(), 1E-8);
	}
	
	/**
	 * source (H11,J13) to destination (G10, L12)
	 */
	@Test
	public void testCutPasteToTop() {
		
		int dstTopRow = 9;
		int dstLeftCol = 6;
		int dstBottomRow = 11;
		int dstRightCol = 11;
		
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange, true);
		
		int realDstRowCount = pasteRange.getLastRow() - pasteRange.getRow();
		int realDstColCount = pasteRange.getLastColumn() - pasteRange.getColumn();
		
		// validate destination
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / SRC_ROW_COUNT, realDstColCount / SRC_COL_COUNT);
		
		// validate source (is clean?)
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 12, 7).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 12, 8).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 12, 9).getCellData().getType().ordinal(), 1E-8);
	}
	
	/**
	 * source (H11,J13) to destination (G10, L15)
	 */
	@Test
	public void testCutPasteToOutsideInclude() {
		
		int dstTopRow = 9;
		int dstLeftCol = 6;
		int dstBottomRow = 14;
		int dstRightCol = 11;
		
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange, true);
		
		int realDstRowCount = pasteRange.getLastRow() - pasteRange.getRow();
		int realDstColCount = pasteRange.getLastColumn() - pasteRange.getColumn();
		
		// validate destination
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / SRC_ROW_COUNT, realDstColCount / SRC_COL_COUNT);
		
		// don't need to validate source. (overlap)
	}
	
	@Test
	public void testCutPasteToAnotherSheet() {
		int dstTopRow = 9;
		int dstLeftCol = 6;
		int dstBottomRow = 11;
		int dstRightCol = 11;
		
		Sheet srcSheet = _workbook.getSheet("Sheet1");
		Sheet dstSheet = _workbook.getSheet("Sheet2");
		Range srcRange = Ranges.range(srcSheet, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(dstSheet, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange, true);
		
		int realDstRowCount = pasteRange.getLastRow() - pasteRange.getRow();
		int realDstColCount = pasteRange.getLastColumn() - pasteRange.getColumn();
		
		// validate destination
		validate(dstSheet, dstTopRow, dstLeftCol, realDstRowCount / SRC_ROW_COUNT, realDstColCount / SRC_COL_COUNT);
		
		// validate source (is clean?)
		for(int i = SRC_TOP_ROW; i <= SRC_BOTTOM_ROW; i++) {
			for(int j = SRC_LEFT_COL; j <= SRC_RIGHT_COL; j++) {
				assertEquals(CellType.BLANK.ordinal(), Ranges.range(srcSheet, i, j).getCellData().getType().ordinal(), 1E-8);
			}
		}
	}
	
	
	/**
	 * validate destination
	 * assume source is always 3 x 3
	 * 1 2 3
	 * 4 5 6
	 * 7 8 9
	 */
	private void validate(Sheet sheet1, int dstTopRow, int dstLeftCol, int rowRepeat, int colRepeat) {
		// validation
		for(int rr = rowRepeat, row = dstTopRow; rr > 0; rr--) {
			for(int cc = colRepeat, col = dstLeftCol; cc > 0; cc--) {
				
				assertEquals(1, Ranges.range(sheet1, row, col).getCellData().getDoubleValue(), 1E-8);
				assertEquals(2, Ranges.range(sheet1, row, col+1).getCellData().getDoubleValue(), 1E-8);
				assertEquals(3, Ranges.range(sheet1, row, col+2).getCellData().getDoubleValue(), 1E-8);
				// next row
				assertEquals(4, Ranges.range(sheet1, row+1, col).getCellData().getDoubleValue(), 1E-8);
				assertEquals(5, Ranges.range(sheet1, row+1, col+1).getCellData().getDoubleValue(), 1E-8);
				assertEquals(6, Ranges.range(sheet1, row+1, col+2).getCellData().getDoubleValue(), 1E-8);
				// next row
				assertEquals(7, Ranges.range(sheet1, row+2, col).getCellData().getDoubleValue(), 1E-8);
				assertEquals(8, Ranges.range(sheet1, row+2, col+1).getCellData().getDoubleValue(), 1E-8);
				assertEquals(9, Ranges.range(sheet1, row+2, col+2).getCellData().getDoubleValue(), 1E-8);
				// move column pointer whenever column repeat
				col += 3;
			}
			// move row pointer whenever row repeat
			row += 3;
		}
	}
	
}
