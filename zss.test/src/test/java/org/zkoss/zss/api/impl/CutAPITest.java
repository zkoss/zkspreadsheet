package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
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
 * 1. destination is correct.
 * 2. source is cleared.
 * summary:
 * 0. simple cut.
 * 1. cut with overlap - 11 overlap state.
 * 2. cut with merge.
 * 3. cut with merge overlap.
 * 4. cut paste to another sheet (undone)
 * @author kuro
 *
 */
public class CutAPITest {
	
	private static Book _workbook;
	// 3 x 3, 1-9 matrix (source)
	// Source Range (H11, J13)
	private static int SRC_TOP_ROW = 10;
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
	 * issue ZSS-401
	 * Cut from one sheet then paste to another sheet will not unmerge source cell on UI.
	 * Server side is clean.
	 * 
	 * ¢x   1    ¢x
	 * ¢x¢w¢w¢w¢w¢w¢w¢w¢w¢x
	 * ¢x 4¢x 5 ¢x6¢x
	 * ¢x 7¢x   ¢x9¢x
	 * 
	 * 1 is a horizontal merged cell 1 x 3.
	 * 4 is a vertical merged cell 2 x 1.
	 * source (H11,J13) sheet1.
	 * destination (I12) sheet2.
	 */
	@Test
	public void testCutPasteRepeatWithMergeIntoAnotherSheet() {
		
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Sheet sheet2 = _workbook.getSheet("Sheet2");
		
		// preparation
		// Merge (H11, J11), horizontal merge, cell value is 1
		Range range_H11J11 = Ranges.range(sheet1, 10, 7, 10, 9);
		range_H11J11.merge(false);
		
		// Merge (I12, I13), vertical merge, cell value is 5
		Range range_I12I13 = Ranges.range(sheet1, 11, 8, 12, 8);
		range_I12I13.merge(false);
		
		// should be merged region
		assertTrue(Util.isAMergedRange(range_H11J11));
		assertTrue(Util.isAMergedRange(range_I12I13));
		
		int dstTopRow = 11;
		int dstLeftCol = 8;
		
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet2, dstTopRow, dstLeftCol);
		Range pasteRange = srcRange.paste(dstRange, true);
		
		int realDstRowCount = pasteRange.getRowCount();;
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / SRC_ROW_COUNT);
		assertEquals(1, realDstColCount / SRC_COL_COUNT);
		
		// validate destination
		
		// this should be a merged cell
		Range horizontalMergedRange = Ranges.range(sheet2, 11, 8, 11, 10);
		assertEquals(1, horizontalMergedRange.getCellData().getDoubleValue(), 1E-8);
		assertTrue(Util.isAMergedRange(horizontalMergedRange));
		
		// this should be a merged cell
		Range verticalMergedRange = Ranges.range(sheet2, 12, 9, 13, 9);
		assertEquals(5, verticalMergedRange.getCellData().getDoubleValue(), 1E-8);
		assertTrue(Util.isAMergedRange(verticalMergedRange));
		
		assertEquals(4, Ranges.range(sheet2, 12, 8).getCellData().getDoubleValue(), 1E-8);
		assertEquals(7, Ranges.range(sheet2, 13, 8).getCellData().getDoubleValue(), 1E-8);
		assertEquals(6, Ranges.range(sheet2, 12, 10).getCellData().getDoubleValue(), 1E-8);
		assertEquals(9, Ranges.range(sheet2, 13, 10).getCellData().getDoubleValue(), 1E-8);
		
		// validate source
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 10, 7).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 10, 8).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 10, 9).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 11, 7).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 11, 8).getCellData().getType().ordinal(), 1E-8);		
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 11, 9).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 12, 7).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 12, 8).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 12, 9).getCellData().getType().ordinal(), 1E-8);

		// should not be merged region anymore
		assertTrue(!Util.isAMergedRange(range_H11J11));
		assertTrue(!Util.isAMergedRange(range_I12I13));
	}
	
	/**
	 * ¢x   1    ¢x
	 * ¢x¢w¢w¢w¢w¢w¢w¢w¢w¢x
	 * ¢x 4¢x 5 ¢x6¢x
	 * ¢x 7¢x   ¢x9¢x
	 * 
	 * 1 is a horizontal merged cell 1 x 3.
	 * 4 is a vertical merged cell 2 x 1.
	 * cut and paste to 5.
	 * source (H11,J13) to destination (I12)
	 */
	@Test
	public void testCutPasteWithMergeOverlap() {
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		
		// preparation
		// Merge (H11, J11), horizontal merge, cell value is 1
		Range range_H11J11 = Ranges.range(sheet1, 10, 7, 10, 9);
		range_H11J11.merge(false);
		
		// Merge (I12, I13), vertical merge, cell value is 5
		Range range_I12I13 = Ranges.range(sheet1, 11, 8, 12, 8);
		range_I12I13.merge(false);
		
		int dstTopRow = 11;
		int dstLeftCol = 8;
		
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol);
		Range pasteRange = srcRange.paste(dstRange, true);
		
		int realDstRowCount = pasteRange.getRowCount();;
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / SRC_ROW_COUNT);
		assertEquals(1, realDstColCount / SRC_COL_COUNT);
		
		// validate destination
		
		// this should be a merged cell
		Range horizontalMergedRange = Ranges.range(sheet1, 11, 8, 11, 10);
		assertEquals(1, horizontalMergedRange.getCellData().getDoubleValue(), 1E-8);
		assertTrue(Util.isAMergedRange(horizontalMergedRange));
		
		// this should be a merged cell
		Range verticalMergedRange = Ranges.range(sheet1, 12, 9, 13, 9);
		assertEquals(5, verticalMergedRange.getCellData().getDoubleValue(), 1E-8);
		assertTrue(Util.isAMergedRange(verticalMergedRange));
		
		assertEquals(4, Ranges.range(sheet1, 12, 8).getCellData().getDoubleValue(), 1E-8);
		assertEquals(7, Ranges.range(sheet1, 13, 8).getCellData().getDoubleValue(), 1E-8);
		assertEquals(6, Ranges.range(sheet1, 12, 10).getCellData().getDoubleValue(), 1E-8);
		assertEquals(9, Ranges.range(sheet1, 13, 10).getCellData().getDoubleValue(), 1E-8);
		
		// validate source
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 10, 7).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 10, 8).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 10, 9).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 11, 7).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 12, 7).getCellData().getType().ordinal(), 1E-8);

		// should not be merged region anymore
		assertTrue(!Util.isAMergedRange(range_H11J11));
		assertTrue(!Util.isAMergedRange(range_I12I13));
	}
	
	/**
	 * ¢x   1    ¢x
	 * ¢x¢w¢w¢w¢w¢w¢w¢w¢w¢x
	 * ¢x 4¢x 5 ¢x6¢x
	 * ¢x 7¢x   ¢x9¢x
	 * 
	 * 1 is a horizontal merged cell 1 x 3.
	 * 4 is a vertical merged cell 2 x 1.
	 * source (H11,J13) to destination (C5, E7)
	 */
	@Test
	public void testCutPasteWithMerge() {
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		
		// preparation
		// Merge (H11, J11), horizontal merge, cell value is 1
		Range range_H11J11 = Ranges.range(sheet1, 10, 7, 10, 9);
		range_H11J11.merge(false);
		
		// Merge (I12, I13), vertical merge, cell value is 5
		Range range_I12I13 = Ranges.range(sheet1, 11, 8, 12, 8);
		range_I12I13.merge(false);
		
		int dstTopRow = 4;
		int dstLeftCol = 2;
		int dstBottomRow = 6;
		int dstRightCol = 4;
		
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange, true);
		
		int realDstRowCount = pasteRange.getRowCount();;
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / SRC_ROW_COUNT);
		assertEquals(1, realDstColCount / SRC_COL_COUNT);
		
		// validate destination
		
		// this should be a merged cell
		Range horizontalMergedRange = Ranges.range(sheet1, 4, 2, 4, 4);
		assertEquals(1, horizontalMergedRange.getCellData().getDoubleValue(), 1E-8);
		assertTrue(Util.isAMergedRange(horizontalMergedRange));
		
		// this should be a merged cell
		Range verticalMergedRange = Ranges.range(sheet1, 5, 3, 6, 3);
		assertEquals(5, verticalMergedRange.getCellData().getDoubleValue(), 1E-8);
		assertTrue(Util.isAMergedRange(verticalMergedRange));
		
		assertEquals(4, Ranges.range(sheet1, 5, 2).getCellData().getDoubleValue(), 1E-8);
		assertEquals(7, Ranges.range(sheet1, 6, 2).getCellData().getDoubleValue(), 1E-8);
		assertEquals(6, Ranges.range(sheet1, 5, 4).getCellData().getDoubleValue(), 1E-8);
		assertEquals(9, Ranges.range(sheet1, 6, 4).getCellData().getDoubleValue(), 1E-8);
		
		// validate source
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL+1).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL+2).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, SRC_TOP_ROW+1, SRC_LEFT_COL).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, SRC_TOP_ROW+1, SRC_LEFT_COL+1).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, SRC_TOP_ROW+1, SRC_LEFT_COL+2).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, SRC_TOP_ROW+2, SRC_LEFT_COL).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, SRC_TOP_ROW+2, SRC_LEFT_COL+1).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, SRC_TOP_ROW+2, SRC_LEFT_COL+2).getCellData().getType().ordinal(), 1E-8);

		// should not be merged region anymore
		assertTrue(!Util.isAMergedRange(range_H11J11));
		assertTrue(!Util.isAMergedRange(range_I12I13));
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
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / SRC_ROW_COUNT);
		assertEquals(1, realDstColCount / SRC_COL_COUNT);		
		
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
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / SRC_ROW_COUNT);
		assertEquals(1, realDstColCount / SRC_COL_COUNT);
		
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
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 10, 7).getCellData().getType().ordinal(), 1E-8);
		
		
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
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(2, realDstRowCount / SRC_ROW_COUNT);
		assertEquals(1, realDstColCount / SRC_COL_COUNT);
		
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
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / SRC_ROW_COUNT);
		assertEquals(2, realDstColCount / SRC_COL_COUNT);
		
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
		
		int realDstRowCount = pasteRange.getRowCount();;
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(2, realDstRowCount / SRC_ROW_COUNT);
		assertEquals(2, realDstColCount / SRC_COL_COUNT);
		
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
		
		int realDstRowCount = pasteRange.getRowCount();;
		int realDstColCount = pasteRange.getColumnCount();
		
		// validate destination
		validate(dstSheet, dstTopRow, dstLeftCol, realDstRowCount / SRC_ROW_COUNT, realDstColCount / SRC_COL_COUNT);
		
		// paste size validation
		assertEquals(1, realDstRowCount / SRC_ROW_COUNT);
		assertEquals(2, realDstColCount / SRC_COL_COUNT);
		
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
