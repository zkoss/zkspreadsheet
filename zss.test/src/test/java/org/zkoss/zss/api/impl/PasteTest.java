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
import org.zkoss.zss.api.IllegalOpArgumentException;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.PasteOperation;
import org.zkoss.zss.api.Range.PasteType;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.CellData.CellType;

/**
 * paste test
 * summary:
 * 0: simple paste.
 * 1: paste to a single cell.
 * 2: paste to a single cell with overlap.
 * 3: repeat paste.
 * 4: repeat paste with overlap.
 * 5: paste with skip blank.
 * 6: paste to another sheet.
 * 7: paste with destination is bigger than source and overlap all the source.
 * 8: paste with transpose.
 * 9: paste with merge.
 * 10: paste with merge and transpose.
 * @author kuro
 */
public class PasteTest {
	
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
	
	@Test(expected = IllegalOpArgumentException.class)
	public void testTransposePasteOverlap2007() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xlsx");
		testTransposePasteOverlap0(book);
	}
	
	@Test(expected = IllegalOpArgumentException.class)
	public void testTransposePasteOverlap2003() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		testTransposePasteOverlap0(book);
	}
	
	@Test 
	public void testPasteMergeTranspose2007() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xlsx");
		testPasteMergeTranspose0(book);
	}
	
	@Test 
	public void testPasteMergeTranspose2003() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		testPasteMergeTranspose0(book);
	}
	
	@Test
	public void testPasteMerge2007() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xlsx");
		testPasteMerge0(book);
	}
	
	@Test
	public void testPasteMerge2003() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		testPasteMerge0(book);
	}
	
	@Test
	public void testPasteToG10_2003() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		testPasteToG10(book);
	}
	
	@Test
	public void testPasteToG10_2007() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xlsx");
		testPasteToG10(book);
	}
	
	@Test
	public void testPasteToG12_2003() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		testPasteToG12(book);
	}
	
	@Test
	public void testPasteToG12_2007() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xlsx");
		testPasteToG12(book);
	}
	
	@Test 
	public void testPasteToI12_2003() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		testPasteToI12(book);
	}
	
	@Test 
	public void testPasteToI12_2007() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		testPasteToI12(book);
	}
	
	@Test
	public void testPasteRepeat2003() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		testPasteRepeat(book);
	}
	
	@Test
	public void testPasteRepeat2007() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xlsx");
		testPasteRepeat(book);
	}
	
	@Test 
	public void testPasteRepeatOverlap2003() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		testPasteRepeatOverlap(book);
	}
	
	@Test 
	public void testPasteRepeatOverlap2007() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xlsx");
		testPasteRepeatOverlap(book);
	}
	
	@Test
	public void testPasteRepeatToAnotherSheet2003() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		testPasteRepeatToAnotherSheet(book);
	}
	
	@Test
	public void testPasteRepeatToAnotherSheet2007() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xlsx");
		testPasteRepeatToAnotherSheet(book);
	}
	
	@Test 
	public void testPasteRepeatOverlapInclude2003() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		testPasteRepeatOverlapInclude(book);
	}
	
	@Test 
	public void testPasteRepeatOverlapInclude2007() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xlsx");
		testPasteRepeatOverlapInclude(book);
	}
	
	@Test
	public void pasteRepeatWithSkipBlankTrue2003() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xls");
		pasteRepeatWithSkipBlankTrue(book);
	}
	
	@Test
	public void pasteRepeatWithSkipBlankTrue2007() throws IOException {
		Book book = Util.loadBook("book/pasteTest.xlsx");
		pasteRepeatWithSkipBlankTrue(book);
	}
	
	
	/**
	 * Should throw exception IllegalOpArgumentException("Cannot transpose paste to overlapped range");
	 */
	private void testTransposePasteOverlap0(Book workbook) {
		
		int srcTopRow = 10;
		int srcLeftCol = 7;
		int srcBottomRow = 12;
		int srcRightCol = 9;
		
		int dstTopRow = 11;
		int dstLeftCol = 8;
		
		// operation
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		// Ranges.range(, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range srcRange = Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol);
		srcRange.pasteSpecial(dstRange, PasteType.ALL, PasteOperation.NONE, false, true);
	}
	
	/**
	 *    1
	 * 4  5  6
	 * 7     9
	 * 
	 * 1 is a horizontal merged cell 1 x 3.
	 * 5 is a vertical merged cell 2 x 1.
	 * transpose paste.
	 * source (H11,J13) to destination (C5, E7).
	 */
	private void testPasteMergeTranspose0(Book workbook) {
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
		Range pasteRange = srcRange.pasteSpecial(dstRange, PasteType.ALL, PasteOperation.NONE, false, true);
		
		int realDstRowCount = pasteRange.getRowCount();;
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / srcRowCount);
		assertEquals(1, realDstColCount / srcColCount);
		
		// validate destination
		
		// this should be a merged cell
		// origin is horizontal, but now vertical
		Range horizontalMergedRange = Ranges.range(sheet1, 4, 2, 6, 2);
		
		assertEquals(1, horizontalMergedRange.getCellData().getDoubleValue(), 1E-8);
		assertTrue(Util.isAMergedRange(horizontalMergedRange));
		
		// this should be a merged cell
		// origin is vertical, but now horizontal
		Range verticalMergedRange = Ranges.range(sheet1, 5, 3, 5, 4);
		assertEquals(5, verticalMergedRange.getCellData().getDoubleValue(), 1E-8);
		assertTrue(Util.isAMergedRange(verticalMergedRange));
		
		assertEquals(4, Ranges.range(sheet1, dstTopRow, dstLeftCol+1).getCellData().getDoubleValue(), 1E-8);
		assertEquals(7, Ranges.range(sheet1, dstTopRow, dstLeftCol+2).getCellData().getDoubleValue(), 1E-8);

		assertEquals(6, Ranges.range(sheet1, dstTopRow+2, dstLeftCol+1).getCellData().getDoubleValue(), 1E-8);
		assertEquals(9, Ranges.range(sheet1, dstTopRow+2, dstLeftCol+2).getCellData().getDoubleValue(), 1E-8);		
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
	private void testPasteMerge0(Book workbook) {
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
		Range pasteRange = srcRange.paste(dstRange);
		
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
		
		assertEquals(4, Ranges.range(sheet1, dstTopRow+1, dstLeftCol).getCellData().getDoubleValue(), 1E-8);
		assertEquals(7, Ranges.range(sheet1, dstTopRow+2, dstLeftCol).getCellData().getDoubleValue(), 1E-8);

		assertEquals(6, Ranges.range(sheet1, dstTopRow+1, dstRightCol).getCellData().getDoubleValue(), 1E-8);
		assertEquals(9, Ranges.range(sheet1, dstTopRow+2, dstRightCol).getCellData().getDoubleValue(), 1E-8);		
	}
	
	private void testPasteToG10(Book workbook) {
		// source (H11, J13)
		int srcTopRow = 10;
		int srcLeftCol = 7;
		int srcBottomRow = 12;
		int srcRightCol = 9;
		int srcRowCount = srcBottomRow - srcTopRow + 1;
		int srcColCount = srcRightCol - srcLeftCol + 1;
		
		// 1 x 1, blank destination
		// dst G10 (on the left top of 1)
		int dstTopRow = 9;
		int dstLeftCol = 6;
		
		// operation
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		// Ranges.range(, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range srcRange = Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol);
		Range pasteRange = srcRange.paste(dstRange); // copy src to dst
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / srcRowCount);
		assertEquals(1, realDstColCount / srcColCount);
		
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / srcRowCount, realDstColCount / srcColCount);
	}
	
	private void testPasteToG12(Book workbook) {
		// source (H11, J13)
		int srcTopRow = 10;
		int srcLeftCol = 7;
		int srcBottomRow = 12;
		int srcRightCol = 9;
		int srcRowCount = srcBottomRow - srcTopRow + 1;
		int srcColCount = srcRightCol - srcLeftCol + 1;
		
		// 1 x 1, blank destination
		// dst G12 (on the left of 4)
		int dstTopRow = 12;
		int dstLeftCol = 6;
		
		// operation
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		// Ranges.range(, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range srcRange = Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol);
		Range pasteRange = srcRange.paste(dstRange); // copy src to dst
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / srcRowCount);
		assertEquals(1, realDstColCount / srcColCount);		
		
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / srcRowCount, realDstColCount / srcColCount);
	}
	
	/**
	 * source (H11, J13) 
	 * destination I12 (on the number 5)
	 * paste to 5
	 */
	private void testPasteToI12(Book workbook) {
		int srcTopRow = 10;
		int srcLeftCol = 7;
		int srcBottomRow = 12;
		int srcRightCol = 9;
		int srcRowCount = srcBottomRow - srcTopRow + 1;
		int srcColCount = srcRightCol - srcLeftCol + 1;
		
		int dstTopRow = 11;
		int dstLeftCol = 8;
		
		// operation
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		// Ranges.range(, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range srcRange = Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol);
		Range pasteRange = srcRange.paste(dstRange); // copy src to dst
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / srcRowCount);
		assertEquals(1, realDstColCount / srcColCount);		
		
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / srcRowCount, realDstColCount / srcColCount);
	}
	
	private void testPasteRepeat(Book workbook) {
		int srcTopRow = 10;
		int srcLeftCol = 7;
		int srcBottomRow = 12;
		int srcRightCol = 9;
		int srcRowCount = srcBottomRow - srcTopRow + 1;
		int srcColCount = srcRightCol - srcLeftCol + 1;
		
		// source (H11, J13) 3 x 3
		// destination (L11, Q13) 6 x 6
		int dstTopRow = 10;
		int dstLeftCol = 11;
		int dstBottomRow = 15;
		int dstRightCol = 16;
		
		// operation
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Range srcRange = Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange); // copy src to dst
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(2, realDstRowCount / srcRowCount);
		assertEquals(2, realDstColCount / srcColCount);		
		
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / srcRowCount, realDstColCount / srcColCount);
	}
	
	/**
	 * paste repeat overlap, left cover
	 * source (H11, J13) 3 x 3
	 * dst (H7, J12) 6 x 3
	 */
	private void testPasteRepeatOverlap(Book workbook) {
		int srcTopRow = 10;
		int srcLeftCol = 7;
		int srcBottomRow = 12;
		int srcRightCol = 9;
		int srcRowCount = srcBottomRow - srcTopRow + 1;
		int srcColCount = srcRightCol - srcLeftCol + 1;

		int dstTopRow = 6;
		int dstLeftCol = 7;
		int dstBottomRow = 11;
		int dstRightCol = 9;
		
		// operation
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Range srcRange = Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange); // copy src to dst
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(2, realDstRowCount / srcRowCount);
		assertEquals(1, realDstColCount / srcColCount);			
		
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / srcRowCount, realDstColCount / srcColCount);
	}
	
	/**
	 * paste source to destination which at different sheet and repeat.
	 * source (H11, J13) sheet1, 3 x 3
	 * dst (G9, L14) sheet2, 6 x 6
	 * 
	 */
	private void testPasteRepeatToAnotherSheet(Book workbook) {
		
		int srcTopRow = 10;
		int srcLeftCol = 7;
		int srcBottomRow = 12;
		int srcRightCol = 9;
		int srcRowCount = srcBottomRow - srcTopRow + 1;
		int srcColCount = srcRightCol - srcLeftCol + 1;

		int dstTopRow = 8;
		int dstLeftCol = 6;
		int dstBottomRow = 13;
		int dstRightCol = 11;
		
		// operation
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Sheet dstSheet = workbook.getSheet("Sheet2");
		Range srcRange = Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range dstRange = Ranges.range(dstSheet, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange); // copy src to dst
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(2, realDstRowCount / srcRowCount);
		assertEquals(2, realDstColCount / srcColCount);		
		
		validate(dstSheet, dstTopRow, dstLeftCol, realDstRowCount / srcRowCount, realDstColCount / srcColCount);		
	}
	
	/**
	 * source (H11, J13) - 3 x 3 block
	 * dst (G9, L14) - 6 x 6 block
	 * destination is bigger than source and include the source inside.
	 */
	private void testPasteRepeatOverlapInclude(Book workbook) {
		int srcTopRow = 10;
		int srcLeftCol = 7;
		int srcBottomRow = 12;
		int srcRightCol = 9;
		int srcRowCount = srcBottomRow - srcTopRow + 1;
		int srcColCount = srcRightCol - srcLeftCol + 1;
		
		int dstTopRow = 8;
		int dstLeftCol = 6;
		int dstBottomRow = 13;
		int dstRightCol = 11;
		
		// operation
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Range srcRange = Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange); // copy src to dst
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(2, realDstRowCount / srcRowCount);
		assertEquals(2, realDstColCount / srcColCount);
		
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / srcRowCount, realDstColCount / srcColCount);
	}
	
	/**
	 * check skip blank is work, blank don't replace the value when paste
	 * in following case: blank doesn't replace the value "1" in K11
	 * B is blank
	 * B 2 3          1 B B
	 * 4 5 6 paste to B B B repeat
	 * 7 8 9          B B B  
	 * result
	 * 1 2 3 B 2 3
	 * 4 5 6 4 5 6
	 * 7 8 9 7 8 9
	 */
	private void pasteRepeatWithSkipBlankTrue(Book workbook) {
		int srcTopRow = 10;
		int srcLeftCol = 7;
		int srcBottomRow = 12;
		int srcRightCol = 9;
		int srcRowCount = srcBottomRow - srcTopRow + 1;
		int srcColCount = srcRightCol - srcLeftCol + 1;
		
		// preparation
		Sheet sheet1 = workbook.getSheet("Sheet1");
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
		Range srcRange = Ranges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.pasteSpecial(dstRange, PasteType.ALL, PasteOperation.NONE, true, false);
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / srcRowCount);
		assertEquals(2, realDstColCount / srcColCount);
		
		// validation
		// is K11 = 1?
		assertEquals(1, cellK11.getCellData().getDoubleValue(), 1E-8);
		// is N11 blank?
		Range cellN11 = Ranges.range(sheet1, 10, 13);
		assertEquals(CellType.BLANK.ordinal(), cellN11.getCellData().getType().ordinal(), 1E-8);
		// fill N11 as 1
		cellN11.setCellValue(1);
		// validate rest
		validate(sheet1, dstTopRow, dstLeftCol, realDstRowCount / srcRowCount, realDstColCount / srcColCount);
		
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
