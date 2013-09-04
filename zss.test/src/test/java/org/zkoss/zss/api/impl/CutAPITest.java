package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.IOException;
import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.Setup;
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
 * 4. cut paste to another sheet (issue ZSS-401)
 * @author kuro
 *
 */
public class CutAPITest {
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		ZssContext.setThreadLocal(new ZssContext(Locale.TAIWAN,-1));
	}
	
	@After
	public void tearDown() throws Exception {
		ZssContext.setThreadLocal(null);
	}
	
	@Test
	public void testCutPasteRepeatWithMergeIntoAnotherSheet2003() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		testCutPasteRepeatWithMergeIntoAnotherSheet0(book);
	}
	
	@Test
	public void testCutPasteRepeatWithMergeIntoAnotherSheet2007() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xlsx");
		testCutPasteRepeatWithMergeIntoAnotherSheet0(book);
	}
	
	@Test
	public void testCutPasteWithMergeOverlap2003() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		testCutPasteWithMergeOverlap0(book);
	}
	
	@Test
	public void testCutPasteWithMergeOverlap2007() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xlsx");
		testCutPasteWithMergeOverlap0(book);
	}
	
	@Test
	public void testCutPasteWithMerge2003() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		testCutPasteWithMerge0(book);
	}
	
	@Test
	public void testCutPasteWithMerge2007() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xlsx");
		testCutPasteWithMerge0(book);
	}
	
	@Test
	public void testCutPasteToLeftTop2003() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		testCutPasteToLeftTop0(book);
	}
	
	@Test
	public void testCutPasteToLeftTop2007() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xlsx");
		testCutPasteToLeftTop0(book);
	}
	
	@Test
	public void testCutPasteToRightBottom2003() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		testCutPasteToRightBottom0(book);
	}
	
	@Test
	public void testCutPasteToRightBottom2007() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xlsx");
		testCutPasteToRightBottom0(book);
	}
	
	@Test
	public void testCutPasteToLeft2003() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		testCutPasteToLeft0(book);
	}
	
	@Test
	public void testCutPasteToLeft2007() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xlsx");
		testCutPasteToLeft0(book);
	}
	
	@Test
	public void testCutPasteToTop2003() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		testCutPasteToTop0(book);
	}
	
	@Test
	public void testCutPasteToTop2007() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xlsx");
		testCutPasteToTop0(book);
	}
	
	@Test
	public void testCutPasteToOutsideInclude2003() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		testCutPasteToOutsideInclude0(book);
	}
	
	@Test
	public void testCutPasteToOutsideInclude2007() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xlsx");
		testCutPasteToOutsideInclude0(book);
	}
	
	@Test
	public void testCutPasteToAnotherSheet2003() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		testCutPasteToAnotherSheet0(book);
	}
	
	@Test
	public void testCutPasteToAnotherSheet2007() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xlsx");
		testCutPasteToAnotherSheet0(book);
	}
	
	/**
	 * relate to issue ZSS-401
	 * Cut from one sheet then paste to another sheet will not unmerge source cell on UI.
	 * Server side is clean.
	 * 
	 *    1
	 * 4  5  6
	 * 7     9
	 * 1 is a horizontal merged cell 1 x 3.
	 * 5 is a vertical merged cell 2 x 1.
	 * source (H11,J13) sheet1.
	 * destination (I12) sheet2.
	 */
	private void testCutPasteRepeatWithMergeIntoAnotherSheet0(Book workbook) {
		
		// 3 x 3, 1-9 matrix (source) 
		// Source Range (H11, J13) Sheet1
		Sheet sheet1 = workbook.getSheet("Sheet1");
		int srcTopRow = 10;
		int srcLeftCol = 7;
		int srcBottomRow = 12;
		int srcRightCol = 9;
		int srcRowCount = srcBottomRow - srcTopRow + 1;
		int srcColCount = srcRightCol - srcLeftCol + 1;
		
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
		
		// destination (I12) sheet2.
		Sheet sheet2 = workbook.getSheet("Sheet2");
		int dstTopRow = 11;
		int dstLeftCol = 8;
		
		Range srcRange = Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range dstRange = Ranges.range(sheet2, dstTopRow, dstLeftCol);
		Range pasteRange = srcRange.paste(dstRange, true);
		
		int realDstRowCount = pasteRange.getRowCount();;
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / srcRowCount);
		assertEquals(1, realDstColCount / srcColCount);
		
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
	 *    1
	 * 4  5  6
	 * 7     9
	 * 
	 * 1 is a horizontal merged cell 1 x 3.
	 * 5 is a vertical merged cell 2 x 1.
	 * cut and paste to 5.
	 * source (H11,J13) to destination (I12)
	 */
	private void testCutPasteWithMergeOverlap0(Book workbook) {
		
		Sheet sheet1 = workbook.getSheet("Sheet1");
		int srcTopRow = 10;
		int srcLeftCol = 7;
		int srcBottomRow = 12;
		int srcRightCol = 9;
		int srcRowCount = srcBottomRow - srcTopRow + 1;
		int srcColCount = srcRightCol - srcLeftCol + 1;
		
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
		
		Range srcRange = Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol);
		Range pasteRange = srcRange.paste(dstRange, true);
		
		int realDstRowCount = pasteRange.getRowCount();;
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / srcRowCount);
		assertEquals(1, realDstColCount / srcColCount);
		
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
	 *    1
	 * 4  5  6
	 * 7     9
	 * 
	 * 1 is a horizontal merged cell 1 x 3.
	 * 5 is a vertical merged cell 2 x 1.
	 * source (H11,J13) to destination (C5, E7)
	 */
	private void testCutPasteWithMerge0(Book workbook) {
		
		Sheet sheet1 = workbook.getSheet("Sheet1");
		int srcTopRow = 10;
		int srcLeftCol = 7;
		int srcBottomRow = 12;
		int srcRightCol = 9;
		int srcRowCount = srcBottomRow - srcTopRow + 1;
		int srcColCount = srcRightCol - srcLeftCol + 1;
		
		
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
		
		int dstTopRow = 4;
		int dstLeftCol = 2;
		int dstBottomRow = 6;
		int dstRightCol = 4;
		
		Range srcRange = Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange, true);
		
		int realDstRowCount = pasteRange.getRowCount();;
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / srcRowCount);
		assertEquals(1, realDstColCount / srcColCount);
		
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
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, srcTopRow, srcLeftCol).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, srcTopRow, srcLeftCol+1).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, srcTopRow, srcLeftCol+2).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, srcTopRow+1, srcLeftCol).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, srcTopRow+1, srcLeftCol+1).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, srcTopRow+1, srcLeftCol+2).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, srcTopRow+2, srcLeftCol).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, srcTopRow+2, srcLeftCol+1).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, srcTopRow+2, srcLeftCol+2).getCellData().getType().ordinal(), 1E-8);

		// should not be merged region anymore
		assertTrue(!Util.isAMergedRange(range_H11J11));
		assertTrue(!Util.isAMergedRange(range_I12I13));
	}
	
	/**
	 * source (H11,J13) to destination (G10, I12)
	 */
	private void testCutPasteToLeftTop0(Book workbook) {

		int srcTopRow = 10;
		int srcLeftCol = 7;
		int srcBottomRow = 12;
		int srcRightCol = 9;
		int srcRowCount = srcBottomRow - srcTopRow + 1;
		int srcColCount = srcRightCol - srcLeftCol + 1;
		
		int dstTopRow = 9;
		int dstLeftCol = 6;
		int dstBottomRow = 11;
		int dstRightCol = 8;
		
		Sheet sheet1 = workbook.getSheet("Sheet1");
		
		Range srcRange = Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange, true);
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / srcRowCount);
		assertEquals(1, realDstColCount / srcColCount);		
		
		// validate destination
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / srcRowCount, realDstColCount / srcColCount);
		
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
	private void testCutPasteToRightBottom0(Book workbook) {

		int srcTopRow = 10;
		int srcLeftCol = 7;
		int srcBottomRow = 12;
		int srcRightCol = 9;
		int srcRowCount = srcBottomRow - srcTopRow + 1;
		int srcColCount = srcRightCol - srcLeftCol + 1;
		
		int dstTopRow = 12;
		int dstLeftCol = 9;
		int dstBottomRow = 14;
		int dstRightCol = 11;
		
		Sheet sheet1 = workbook.getSheet("Sheet1");
		
		Range srcRange = Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange, true);
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / srcRowCount);
		assertEquals(1, realDstColCount / srcColCount);
		
		// validate destination
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / srcRowCount, realDstColCount / srcColCount);
		
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
	private void testCutPasteToLeft0(Book workbook) {
		
		int dstTopRow = 8;
		int dstLeftCol = 6;
		int dstBottomRow = 13;
		int dstRightCol = 8;
		
		int srcTopRow = 10;
		int srcLeftCol = 7;
		int srcBottomRow = 12;
		int srcRightCol = 9;
		int srcRowCount = srcBottomRow - srcTopRow + 1;
		int srcColCount = srcRightCol - srcLeftCol + 1;
		
		Sheet sheet1 = workbook.getSheet("Sheet1");
		
		Range srcRange = Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange, true);
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(2, realDstRowCount / srcRowCount);
		assertEquals(1, realDstColCount / srcColCount);
		
		// validate destination
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / srcRowCount, realDstColCount / srcColCount);
		
		// validate source (is clean?)
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 10, 9).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 11, 9).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 12, 9).getCellData().getType().ordinal(), 1E-8);
	}
	
	/**
	 * source (H11,J13) to destination (G10, L12)
	 */
	private void testCutPasteToTop0(Book workbook) {
		
		// 3 x 3, 1-9 matrix (source) 
		// Source Range (H11, J13) Sheet1
		Sheet sheet1 = workbook.getSheet("Sheet1");
		int srcTopRow = 10;
		int srcLeftCol = 7;
		int srcBottomRow = 12;
		int srcRightCol = 9;
		int srcRowCount = srcBottomRow - srcTopRow + 1;
		int srcColCount = srcRightCol - srcLeftCol + 1;
		
		int dstTopRow = 9;
		int dstLeftCol = 6;
		int dstBottomRow = 11;
		int dstRightCol = 11;
		
		Range srcRange = Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange, true);
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / srcRowCount);
		assertEquals(2, realDstColCount / srcColCount);
		
		// validate destination
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / srcRowCount, realDstColCount / srcColCount);
		
		// validate source (is clean?)
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 12, 7).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 12, 8).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, 12, 9).getCellData().getType().ordinal(), 1E-8);
	}
	
	/**
	 * source (H11,J13) to destination (G10, L15)
	 */
	private void testCutPasteToOutsideInclude0(Book workbook) {
		
		int srcTopRow = 10;
		int srcLeftCol = 7;
		int srcBottomRow = 12;
		int srcRightCol = 9;
		int srcRowCount = srcBottomRow - srcTopRow + 1;
		int srcColCount = srcRightCol - srcLeftCol + 1;
		
		int dstTopRow = 9;
		int dstLeftCol = 6;
		int dstBottomRow = 14;
		int dstRightCol = 11;
		
		Sheet sheet1 = workbook.getSheet("Sheet1");
		
		Range srcRange = Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange, true);
		
		int realDstRowCount = pasteRange.getRowCount();;
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(2, realDstRowCount / srcRowCount);
		assertEquals(2, realDstColCount / srcColCount);
		
		// validate destination
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / srcRowCount, realDstColCount / srcColCount);
		
		// don't need to validate source. (overlap)
	}
	
	private void testCutPasteToAnotherSheet0(Book workbook) {
		
		int srcTopRow = 10;
		int srcLeftCol = 7;
		int srcBottomRow = 12;
		int srcRightCol = 9;
		int srcRowCount = srcBottomRow - srcTopRow + 1;
		int srcColCount = srcRightCol - srcLeftCol + 1;
		
		int dstTopRow = 9;
		int dstLeftCol = 6;
		int dstBottomRow = 11;
		int dstRightCol = 11;
		
		Sheet srcSheet = workbook.getSheet("Sheet1");
		Sheet dstSheet = workbook.getSheet("Sheet2");
		
		Range srcRange = Ranges.range(srcSheet, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range dstRange = Ranges.range(dstSheet, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange, true);
		
		int realDstRowCount = pasteRange.getRowCount();;
		int realDstColCount = pasteRange.getColumnCount();
		
		// validate destination
		validate(dstSheet, dstTopRow, dstLeftCol, realDstRowCount / srcRowCount, realDstColCount / srcColCount);
		
		// paste size validation
		assertEquals(1, realDstRowCount / srcRowCount);
		assertEquals(2, realDstColCount / srcColCount);
		
		// validate source (is clean?)
		for(int i = srcTopRow; i <= srcBottomRow; i++) {
			for(int j = srcLeftCol; j <= srcRightCol; j++) {
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
