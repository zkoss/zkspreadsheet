package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;

import java.io.InputStream;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.PasteOperation;
import org.zkoss.zss.api.Range.PasteType;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.CellData.CellType;

/**
 * paste overlap test
 * @author kuro
 */
public class PasteOverlapTest {
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
		final String filename = "book/overlappingPaste.xlsx";
		final InputStream is = PasteOverlapTest.class.getResourceAsStream(filename);
		_workbook = Importers.getImporter().imports(is, filename);
	}
	
	@After
	public void tearDown() throws Exception {
		_workbook = null;
	}
	
	@Test
	public void testPasteToG10() {
		// source (H11, J13)
		// 1 x 1, blank destination
		// dst G10 (on the left top of 1)
		int dstTopRow = 9;
		int dstLeftCol = 6;
		
		// operation
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		// Ranges.range(, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol);
		Range pasteRange = srcRange.paste(dstRange); // copy src to dst
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / SRC_ROW_COUNT, realDstColCount / SRC_COL_COUNT);
	}
	
	@Test
	public void testPasteToG12() {
		// source (H11, J13)
		// 1 x 1, blank destination
		// dst G12 (on the left of 4)
		int dstTopRow = 12;
		int dstLeftCol = 6;
		
		// operation
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		// Ranges.range(, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol);
		Range pasteRange = srcRange.paste(dstRange); // copy src to dst
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / SRC_ROW_COUNT, realDstColCount / SRC_COL_COUNT);
	}
	
	@Test 
	public void testPasteToI12() {
		// source (H11, J13)
		// 1 x 1, blank destination
		// dst I12 (on the number 5)
		int dstTopRow = 11;
		int dstLeftCol = 8;
		
		// operation
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		// Ranges.range(, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol);
		Range pasteRange = srcRange.paste(dstRange); // copy src to dst
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / SRC_ROW_COUNT, realDstColCount / SRC_COL_COUNT);
	}
	
	@Test
	public void testPasteRepeat() {
		// source (H11, J13)
		// destination (L11, Q13)
		int dstTopRow = 10;
		int dstLeftCol = 11;
		int dstBottomRow = 16;
		int dstRightCol = 12;
		
		// operation
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange); // copy src to dst
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / SRC_ROW_COUNT, realDstColCount / SRC_COL_COUNT);
	}
	
	@Test 
	public void testPasteRepeatOverlap() {
		// source (H11, J13)
		// 1 x 1, blank destination
		// dst (H7, J12), left cover
		int dstTopRow = 6;
		int dstLeftCol = 7;
		int dstBottomRow = 11;
		int dstRightCol = 9;
		
		// operation
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange); // copy src to dst
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / SRC_ROW_COUNT, realDstColCount / SRC_COL_COUNT);
	}
	
	/**
	 * paste source to destination which at different sheet and repeat.
	 */
	@Test
	public void testPasteRepeatToAnotherSheet() {
		// source (H11, J13) sheet1
		// 1 x 1, blank destination
		// dst (G9, L14) sheet2
		int dstTopRow = 8;
		int dstLeftCol = 6;
		int dstBottomRow = 13;
		int dstRightCol = 11;
		
		// operation
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Sheet dstSheet = _workbook.getSheet("Sheet2");
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(dstSheet, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange); // copy src to dst
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		
		validate(dstSheet, dstTopRow, dstLeftCol, realDstRowCount / SRC_ROW_COUNT, realDstColCount / SRC_COL_COUNT);		
	}	
	
	/**
	 * source (H11, J13) - 3 x 3 block
	 * dst (G9, L14) - 6 x 6 block
	 * destination is bigger than source and include the source inside.
	 */
	@Test 
	public void testPasteRepeatOverlapInclude() {
		int dstTopRow = 8;
		int dstLeftCol = 6;
		int dstBottomRow = 13;
		int dstRightCol = 11;
		
		// operation
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange); // copy src to dst
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		assertEquals(2, realDstRowCount / SRC_ROW_COUNT);
		assertEquals(2, realDstColCount / SRC_COL_COUNT);
		
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / SRC_ROW_COUNT, realDstColCount / SRC_COL_COUNT);
	}
	
	/**
	 * check skip blank is work, blank don't replace the value when paste
	 * in following case: blank doesn't replace the value "1" in K11
	 *   2 3       1
	 * 4 5 6 paste 
	 * 7 8 9
	 * result
	 * 1 2 3   2 3
	 * 4 5 6 4 5 6
	 * 7 8 9 7 8 9
	 */
	@Test
	public void pasteRepeatWithSkipBlankTrue() {
		
		// preparation
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		// let H11 = Blank
		Range cellH11 = Ranges.range(sheet1, 10, 7);
		cellH11.clearContents();
		// let K11 = 1
		Range cellK11 = Ranges.range(sheet1, 10, 10);
		cellK11.setCellValue(1);
		
		// operation
		int dstTopRow = 10;
		int dstLeftCol = 10;
		int dstBottomRow = 12;
		int dstRightCol = 15;		
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.pasteSpecial(dstRange, PasteType.ALL, PasteOperation.NONE, true, false);
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		// validation
		// is K11 = 1?
		assertEquals(1, cellK11.getCellData().getDoubleValue(), 1E-8);
		// is N11 blank?
		Range cellN11 = Ranges.range(sheet1, 13, 10);
		assertEquals(CellType.BLANK.ordinal(), cellN11.getCellData().getType().ordinal(), 1E-8);
		// fill N11 as 1
		cellN11.setCellValue(1);
		// validate rest
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / SRC_ROW_COUNT, realDstColCount / SRC_COL_COUNT);
		
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
