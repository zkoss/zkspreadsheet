package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.IOException;
import java.io.InputStream;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.formula.eval.NotImplementedException;
import org.zkoss.poi.ss.usermodel.ErrorConstants;
import org.zkoss.zss.Setup;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.InsertShift;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.Range.DeleteShift;
import org.zkoss.zss.api.Range.InsertCopyOrigin;
import org.zkoss.zss.api.Range.PasteOperation;
import org.zkoss.zss.api.Range.PasteType;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.CellData.CellType;
import org.zkoss.zss.model.sys.XBook;

/**
 * ZSS-261.
 * ZSS-266.
 * ZSS-275.
 * ZSS-277.
 * ZSS-290.
 * ZSS-298.
 * ZSS-300.
 * ZSS-301.
 * ZSS-303.
 * ZSS-315.
 * ZSS-326. (UNDONE)
 * ZSS-389.
 * ZSS-395.
 */
public class Issue200Test {
	
	private static Book _workbook;
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@After
	public void tearDown() throws Exception {
		_workbook = null;
	}
	
	/**
	 * SERIESSUM(B105,0,2,B106:B109) is a unsupported function, cell value should be ERROR_NAME
	 */
	public void testZSS261() throws IOException {
		final String filename = "book/math.xlsx";
		final InputStream is = PasteTest.class.getResourceAsStream(filename);
		_workbook = Importers.getImporter().imports(is, filename);
		Sheet sheet1 = _workbook.getSheet("formula-math");
		Range B104 = Ranges.range(sheet1, "B104");
		assertEquals(ErrorConstants.ERROR_NAME, B104.getCellData().getDoubleValue(), 1E-8);
	}
	
	@Test
	public void testZSS266() throws IOException {
		final String filename = "book/266-info-formula.xlsx";
		final InputStream is = PasteTest.class.getResourceAsStream(filename);
		_workbook = Importers.getImporter().imports(is, filename);
		Sheet sheet1 = _workbook.getSheet("formula-info");
		Range B12 = Ranges.range(sheet1, "B12");
		assertTrue(B12.getCellData().getBooleanValue());
	}
	
	/**
	 * if a sheet contains "data validation" configuration with empty criteria, 
	 * it causes null pointer exception when displaying it, then crashed.
	 */
	@Test
	public void testZSS255() throws IOException {
		final String filename = "book/255-cell-data.xlsx";
		final InputStream is = PasteTest.class.getResourceAsStream(filename);
		_workbook = Importers.getImporter().imports(is, filename);
	}
	
	@Test
	public void testZSS260() throws IOException {
		final String filename = "book/260-validation.xlsx";
		final InputStream is = PasteTest.class.getResourceAsStream(filename);
		_workbook = Importers.getImporter().imports(is, filename);
		Sheet sheet1 = _workbook.getSheet("Validation");
		Range C3 = Ranges.range(sheet1, "C3");
		C3.setCellEditText("2");
		
		Range C5 = Ranges.range(sheet1, "C5");
		C5.setCellEditText("2013/1/2");
	}
	
	/**
	 * Range.getRow() return unexpected row number
	 */
	@Test
	public void testZSS275() {
		loadPasteTest();
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Range rangeA = Ranges.range(sheet1, "H11:J13");
		assertEquals(10, rangeA.getRow());
	}
	
	/**
	 * 1. merge a 3 x 3 rangeA
	 * 2. create a whole row rangeB which cross the merged cell
	 * 3. perform unmerge on the rangeB
	 * 4. rangeA should be unmerged
	 */
	@Test
	public void testZSS395_1() {
		loadPasteTest();
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Range rangeA = Ranges.range(sheet1, "H11:J13");
		rangeA.merge(false); // merge a 3 x 3
		assertTrue(Util.isAMergedRange(rangeA)); // it should be merged
		Range rangeB = Ranges.range(sheet1, "H12"); // a whole row cross the merged cell
		rangeB.toRowRange().unmerge(); // perform unmerge operation
		assertTrue(!Util.isAMergedRange(rangeA)); // should be unmerged now
	}
	
	/**
	 * 1. merge a 3 x 3 rangeA
	 * 2. create a whole column rangeB which cross the merged cell
	 * 3. perform unmerge on the rangeB
	 * 4. rangeA should be unmerged
	 */
	@Test
	public void testZSS395_2() {
		loadPasteTest();
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Range rangeA = Ranges.range(sheet1, "H11:J13");
		rangeA.merge(false); // merge a 3 x 3
		assertTrue(Util.isAMergedRange(rangeA)); // it should be merged
		Range rangeB = Ranges.range(sheet1, "H12"); // a whole row cross the merged cell
		rangeB.toColumnRange().unmerge(); // perform unmerge operation
		assertTrue(!Util.isAMergedRange(rangeA)); // should be unmerged now
	}
	
	/**
	 * cut a merged cell and paste to another cell, the original cell doesn't become unmerged cells
	 * �x   1    �x
	 * �x�w�w�w�w�w�w�w�w�x
	 * �x 4�x 5 �x6�x
	 * �x 7�x   �x9�x
	 * 
	 * 1 is a horizontal merged cell 1 x 3.
	 * 4 is a vertical merged cell 2 x 1.
	 * source (H11,J13) to destination (C5, E7)
	 */
	@Test
	public void testZSS301() {
		
		loadPasteTest();
		
		int SRC_TOP_ROW = 10;
		int SRC_LEFT_COL = 7;
		int SRC_BOTTOM_ROW = 12;
		int SRC_RIGHT_COL = 9;
		int SRC_ROW_COUNT = SRC_BOTTOM_ROW - SRC_TOP_ROW + 1;
		int SRC_COL_COUNT = SRC_RIGHT_COL - SRC_LEFT_COL + 1;
		
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		
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
	 * Delete multiple columns causes a exception
	 * Delete column E F
	 */
	@Test
	public void testZSS303() {
		
		loadShiftTest();
		
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Range range_E = Ranges.range(sheet1, "E3:F3");
		range_E.toColumnRange().delete(DeleteShift.DEFAULT);
			
		assertEquals("G3", Ranges.range(sheet1, "E3").getCellData().getEditText());
		assertEquals("G4", Ranges.range(sheet1, "E4").getCellData().getEditText());
		assertEquals("G5", Ranges.range(sheet1, "E5").getCellData().getEditText());
		assertEquals("G8", Ranges.range(sheet1, "E8").getCellData().getEditText());
		
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E6").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E7").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "F3").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "F4").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "F5").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "F6").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "F7").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "F8").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G3").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G4").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G5").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G6").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G7").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G8").getCellData().getType().ordinal(), 1E-8);
	}
	
	/**
	 * perform "clear style" on a merged cell doesn't split it back to original unmerged cells
	 */
	@Test
	public void testZSS298() {
		loadPasteTest();
		Sheet sheet = _workbook.getSheet("Sheet1");
		Range range = Ranges.range(sheet, "H11:J13");
		range.merge(false); // 1. merge
		assertTrue(Util.isAMergedRange(range)); // is it merged?
		CellOperationUtil.clearStyles(range);
		assertTrue(!Util.isAMergedRange(range)); // is it unmerged?
	}
	
	@Test
	public void testZSS326() {
		// @FIXME
		/*
		loadZSS326();
		Sheet sheet = _workbook.getSheet("chart-image");
		Range range = Ranges.range(sheet, "B5:B19");
		ChartData cd = ChartDataUtil.getChartData(sheet, new Rect(range.getRow(), range.getColumn(), range.getLastRow(), range.getLastColumn()), Chart.Type.LINE);
		SheetOperationUtil.addChart(range, cd, Chart.Type.LINE, Grouping.STANDARD, LegendPosition.TOP);
		*/
	}
	
	/**
	 * after merging and unmerging cells, input and style both work incorrect
	 */
	@Test
	public void testZSS290() {
		loadZSS290();
		Sheet sheet = _workbook.getSheet("cell");
		Range range = Ranges.range(sheet, "B5:D7");
		range.merge(true);
		range.merge(false);
		range.unmerge();
		String color = "#548dd4";
		assertEquals(color, Ranges.range(sheet, "B5").getCellStyle().getBackgroundColor().getHtmlColor());
		assertEquals(color, Ranges.range(sheet, "B6").getCellStyle().getBackgroundColor().getHtmlColor());
		assertEquals(color, Ranges.range(sheet, "B7").getCellStyle().getBackgroundColor().getHtmlColor());
		assertEquals(color, Ranges.range(sheet, "C5").getCellStyle().getBackgroundColor().getHtmlColor());
		assertEquals(color, Ranges.range(sheet, "C6").getCellStyle().getBackgroundColor().getHtmlColor());
		assertEquals(color, Ranges.range(sheet, "C7").getCellStyle().getBackgroundColor().getHtmlColor());
		assertEquals(color, Ranges.range(sheet, "D5").getCellStyle().getBackgroundColor().getHtmlColor());
		assertEquals(color, Ranges.range(sheet, "D6").getCellStyle().getBackgroundColor().getHtmlColor());
		assertEquals(color, Ranges.range(sheet, "D7").getCellStyle().getBackgroundColor().getHtmlColor());
	}
	
	/**
	 * �x   1    �x
	 * �x�w�w�w�w�w�w�w�w�x
	 * �x 4�x 5 �x6�x
	 * �x 7�x   �x9�x
	 * 
	 * 1 is a horizontal merged cell 1 x 3.
	 * 4 is a vertical merged cell 2 x 1.
	 * transpose paste.
	 * source (H11,J13) to destination (C5, E7).
	 */
	@Test 
	public void testZSS277() {
		
		loadPasteTest();
		
		int SRC_TOP_ROW = 10;
		int SRC_LEFT_COL = 7;
		int SRC_BOTTOM_ROW = 12;
		int SRC_RIGHT_COL = 9;
		int SRC_ROW_COUNT = SRC_BOTTOM_ROW - SRC_TOP_ROW + 1;
		int SRC_COL_COUNT = SRC_RIGHT_COL - SRC_LEFT_COL + 1;
		
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		
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
		
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.pasteSpecial(dstRange, PasteType.ALL, PasteOperation.NONE, false, true);
		
		int realDstRowCount = pasteRange.getRowCount();;
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / SRC_ROW_COUNT);
		assertEquals(1, realDstColCount / SRC_COL_COUNT);
		
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
	 * paste repeat overlap, left cover
	 * source (H11, J13) 3 x 3
	 * dst (H7, J12) 6 x 3
	 */
	@Test 
	public void testPasteZSS300() {
		
		loadPasteTest();

		int SRC_TOP_ROW = 10;
		int SRC_LEFT_COL = 7;
		int SRC_BOTTOM_ROW = 12;
		int SRC_RIGHT_COL = 9;
		int SRC_ROW_COUNT = SRC_BOTTOM_ROW - SRC_TOP_ROW + 1;
		int SRC_COL_COUNT = SRC_RIGHT_COL - SRC_LEFT_COL + 1;
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
		
		// paste size validation
		assertEquals(2, realDstRowCount / SRC_ROW_COUNT);
		assertEquals(1, realDstColCount / SRC_COL_COUNT);			
		
		validateZSS300(sheet1, dstTopRow, dstLeftCol, realDstRowCount / SRC_ROW_COUNT, realDstColCount / SRC_COL_COUNT);
	}
	
	/**
	 * E3:E5 shift up
	 */
	@Test
	public void testZSS389_1() {
		
		loadShiftTest();
		
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Range range_E3E5 = Ranges.range(sheet1, "E3:E5");
		range_E3E5.delete(DeleteShift.UP);
		
		// validate E1, E2 is still correct
		assertEquals("E1", Ranges.range(sheet1, "E1").getCellData().getEditText());
		assertEquals("E2", Ranges.range(sheet1, "E2").getCellData().getEditText());
		
		// validate shift up
		assertEquals("E6", Ranges.range(sheet1, "E3").getCellData().getEditText());
		assertEquals("E7", Ranges.range(sheet1, "E4").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet1, "E5").getCellData().getEditText());
		
		// validate E6, E7, E8 is blank
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E6").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E7").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E8").getCellData().getType().ordinal(), 1E-8);
	}
	
	/**
	 * G3:G5 shift up
	 */
	@Test
	public void testZSS389_2() {
		
		loadShiftTest();
		
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Range range_G3G5 = Ranges.range(sheet1, "G3:G5");
		range_G3G5.delete(DeleteShift.UP);
		
		// validate shift up
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G3").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G4").getCellData().getType().ordinal(), 1E-8);
		assertEquals("G8", Ranges.range(sheet1, "G5").getCellData().getEditText());
		
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G6").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G7").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G8").getCellData().getType().ordinal(), 1E-8);
	}
	
	/**
	 * �x 1 2  3 �x
	 * �x�w�w�w�w�w�w�w�w�x
	 * �x   4    |
	 * |�w�w�w�w�w�w�w�w�x 
	 * �x 7 8  9 �x
	 * 
	 * 4 is a horizontal merged cell 1 x 3.
	 * paste 1 to 4.
	 * source (H11) to destination (H12).
	 * a single cell paste to a merged cell.
	 */
	@Test
	public void testZSS315_1() {
		
		loadPasteTest();
		
		int SRC_TOP_ROW = 10;
		int SRC_LEFT_COL = 7;
		
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		
		// preparation
		// Merge (H12, J12), horizontal merge, cell value is 4
		Range range_H12J12 = Ranges.range(sheet1, 11, 7, 11, 9);
		range_H12J12.merge(false);
		
		assertEquals(4, range_H12J12.getCellData().getDoubleValue(), 1E-8);
		
		// should be merged region
		assertTrue(Util.isAMergedRange(range_H12J12));
		
		int dstTopRow = 11;
		int dstLeftCol = 7;
		
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol);
		Range pasteRange = srcRange.paste(dstRange);
		
		int realDstRowCount = pasteRange.getRowCount();;
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / 1);
		assertEquals(1, realDstColCount / 1);
		
		// this should still be a merged cell
		assertTrue(Util.isAMergedRange(range_H12J12));
		
		// value validation
		assertEquals(1, pasteRange.getCellData().getDoubleValue(), 1E-8);		
	}
	
	/**
	 * �x   1    �x
	 * �x�w�w�w�w�w�w�w�w�x
	 * �x 4 5 6  |
	 * |�w�w�w�w�w�w�w�w�x 
	 * �x 7 8  9 �x
	 * 
	 * 1 is a horizontal merged cell 1 x 3.
	 * paste 1 to 4.
	 * source (H11) to destination (H12).
	 * a single merged cell paste to a single cell.
	 */
	@Test
	public void testZSS315_2() {
		
		loadPasteTest();

		int SRC_TOP_ROW = 10;
		int SRC_LEFT_COL = 7;
		int SRC_RIGHT_COL = 9;
		
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		
		// preparation
		// Merge (H11, J11), horizontal merge, cell value is 1
		Range range_H11J11 = Ranges.range(sheet1, 10, 7, 10, 9);
		range_H11J11.merge(false);
		
		assertEquals(1, range_H11J11.getCellData().getDoubleValue(), 1E-8);
		
		// should be merged region
		assertTrue(Util.isAMergedRange(range_H11J11));
		
		int dstTopRow = 11;
		int dstLeftCol = 7;
		
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_TOP_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol);
		Range pasteRange = srcRange.paste(dstRange);
		
		int realDstRowCount = pasteRange.getRowCount();;
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / 1);
		assertEquals(1, realDstColCount / 3);
		
		// this should be a merged cell
		assertTrue(Util.isAMergedRange(pasteRange));
		
		// value validation
		assertEquals(1, pasteRange.getCellData().getDoubleValue(), 1E-8);
		
	}
	
	private void loadPasteTest() {
		final String filename = "book/pasteTest.xlsx";
		final InputStream is = PasteTest.class.getResourceAsStream(filename);
		try {
			_workbook = Importers.getImporter().imports(is, filename);
		} catch (IOException e) {
			e.printStackTrace();
		}		
	}
	
	private void loadShiftTest() {
		final String filename = "book/shiftTest.xlsx";
		final InputStream is = PasteTest.class.getResourceAsStream(filename);
		try {
			_workbook = Importers.getImporter().imports(is, filename);
		} catch (IOException e) {
			e.printStackTrace();
		}	
	}
	
	private void loadZSS290() {
		final String filename = "book/290-merge.xls";
		final InputStream is = PasteTest.class.getResourceAsStream(filename);
		try {
			_workbook = Importers.getImporter().imports(is, filename);
		} catch (IOException e) {
			e.printStackTrace();
		}	
	}
	
	private void loadZSS326() {
		final String filename = "book/insert-charts.xlsx";
		final InputStream is = PasteTest.class.getResourceAsStream(filename);
		try {
			_workbook = Importers.getImporter().imports(is, filename);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	private void validateZSS300(Sheet sheet1, int dstTopRow, int dstLeftCol, int rowRepeat, int colRepeat) {
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
	
	
	//not fix yet
	@Test
	public void testZSS179() throws IOException {

		final String filename = "book/179-insertexception-simple.xlsx";//should also test non-simple one
		final InputStream is = getClass().getResourceAsStream(filename);
		Book book = Importers.getImporter().imports(is, filename);

		
		Range r = Ranges.range(book.getSheetAt(0), "A1");
    	Assert.assertEquals("A", r.getCellEditText());
    	r = Ranges.range(book.getSheetAt(0), "B1");
    	Assert.assertEquals("", r.getCellEditText());
    	r = Ranges.range(book.getSheetAt(0), "C1");
    	Assert.assertEquals("B", r.getCellEditText());
		
		r = Ranges.range(book.getSheetAt(0), "B1");
		r.insert(InsertShift.RIGHT, InsertCopyOrigin.FORMAT_LEFT_ABOVE);
		
		// export work book
    	Exporter exporter = Exporters.getExporter();
		java.io.File temp = java.io.File.createTempFile("test",".xlsx");
    	java.io.FileOutputStream fos = new java.io.FileOutputStream(temp);
    	//export first time
    	exporter.export(book, fos);
    	r = Ranges.range(book.getSheetAt(0), "A1");
    	Assert.assertEquals("A", r.getCellEditText());
    	r = Ranges.range(book.getSheetAt(0), "B1");
    	Assert.assertEquals("", r.getCellEditText());
    	r = Ranges.range(book.getSheetAt(0), "C1");
    	Assert.assertEquals("", r.getCellEditText());
    	r = Ranges.range(book.getSheetAt(0), "D1");
    	Assert.assertEquals("B", r.getCellEditText());
		
	}
	
}
