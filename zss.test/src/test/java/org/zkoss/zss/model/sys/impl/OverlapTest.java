package org.zkoss.zss.model.sys.impl;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.InputStream;
import java.util.List;
import java.util.Set;

import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.xssf.usermodel.XSSFWorkbook;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.api.model.CellData.CellType;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.XRanges;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.model.sys.impl.XRangeImpl.OverlapState;

/**
 * @author kuro
 *
 */
public class OverlapTest {
	
	private static XBook _workbook;
	// 3 x 3, 1-9 matrix (source)
	private static 	int srcTopRow = 10;
	private static int srcLeftCol = 7;
	private static int srcBottomRow = 12;
	private static int srcRightCol = 9;
	
	@Before
	public void setUp() throws Exception {
		final String filename = "overlappingPaste.xlsx";
		final InputStream is = new ClassLocator().getResourceAsStream(filename);
		_workbook = new ExcelImporter().imports(is, filename);
	}
	
	@After
	public void tearDown() throws Exception {
		_workbook = null;
	}
	
	@Test
	public void testCutPasteToLeftTop() {
		int dstTopRow = 9;
		int dstLeftCol = 6;
		int dstBottomRow = 11;
		int dstRightCol = 8;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		srcRange.copy(dstRange, true);
		
		// validate destination
		assertEquals(1, sheet1.getRow(dstTopRow).getCell(dstLeftCol).getNumericCellValue(), 1E-8);
		assertEquals(2, sheet1.getRow(dstTopRow).getCell(dstLeftCol+1).getNumericCellValue(), 1E-8);
		assertEquals(3, sheet1.getRow(dstTopRow).getCell(dstLeftCol+2).getNumericCellValue(), 1E-8);
		assertEquals(4, sheet1.getRow(++dstTopRow).getCell(dstLeftCol).getNumericCellValue(), 1E-8);
		assertEquals(5, sheet1.getRow(dstTopRow).getCell(dstLeftCol+1).getNumericCellValue(), 1E-8);
		assertEquals(6, sheet1.getRow(dstTopRow).getCell(dstLeftCol+2).getNumericCellValue(), 1E-8);
		assertEquals(7, sheet1.getRow(++dstTopRow).getCell(dstLeftCol).getNumericCellValue(), 1E-8);
		assertEquals(8, sheet1.getRow(dstTopRow).getCell(dstLeftCol+1).getNumericCellValue(), 1E-8);
		assertEquals(9, sheet1.getRow(dstTopRow).getCell(dstLeftCol+2).getNumericCellValue(), 1E-8);
		
		// validate source (is clean?)
		assertEquals(CellType.BLANK.ordinal(), sheet1.getRow(10).getCell(9).getCellType(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), sheet1.getRow(11).getCell(9).getCellType(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), sheet1.getRow(12).getCell(7).getCellType(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), sheet1.getRow(12).getCell(8).getCellType(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), sheet1.getRow(12).getCell(9).getCellType(), 1E-8);
		
		
	}
	
	@Test
	public void testNotOverlap() {
		
		int dstTopRow = 20;
		int dstLeftCol = 15;
		int dstBottomRow = 24;
		int dstRightCol = 18;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		int result = new XRangeImpl(0, 0, sheet1, sheet1).getOverlapState(srcRef, dstRef).ordinal();
		assertEquals(OverlapState.NOT_OVERLAP.ordinal(), result, 1E-8);
	}
	
	@Test
	public void testRef1InsideRef2() {
		
		int dstTopRow = 9;
		int dstLeftCol = 6;
		int dstBottomRow = 13;
		int dstRightCol = 10;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		int result = new XRangeImpl(0, 0, sheet1, sheet1).getOverlapState(srcRef, dstRef).ordinal();
		assertEquals(OverlapState.INCLUDED_OUTSIDE.ordinal(), result, 1E-8);
	}
	
	@Test
	public void testRef2InsideRef1() {
		
		int dstTopRow = 10;
		int dstLeftCol = 7;
		int dstBottomRow = 11;
		int dstRightCol = 8;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		int result = new XRangeImpl(0, 0, sheet1, sheet1).getOverlapState(srcRef, dstRef).ordinal();
		assertEquals(OverlapState.INCLUDED_INSIDE.ordinal(), result, 1E-8);
	}

	@Test
	public void testOverlapLeftTop() {
		
		int dstTopRow = 9;
		int dstLeftCol = 6;
		int dstBottomRow = 11;
		int dstRightCol = 8;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		int result = new XRangeImpl(0, 0, sheet1, sheet1).getOverlapState(srcRef, dstRef).ordinal();
		assertEquals(OverlapState.ON_LEFT_TOP.ordinal(), result, 1E-8);
	}
	
	@Test
	public void testIntersectLeftTop() {
		int dstTopRow = 9;
		int dstLeftCol = 6;
		int dstBottomRow = 11;
		int dstRightCol = 8;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		Ref resultRef = new XRangeImpl(0, 0, sheet1, sheet1).intersect(srcRef, dstRef);
		
		final int topRow = resultRef.getTopRow();
		final int bottomRow = resultRef.getBottomRow();
		final int leftCol = resultRef.getLeftCol();
		final int rightCol = resultRef.getRightCol();
		
		assertEquals(10, topRow, 1E-8);
		assertEquals(7, leftCol, 1E-8);
		assertEquals(11, bottomRow, 1E-8);
		assertEquals(8, rightCol, 1E-8);
	}
	
	@Test
	public void testComplmentLeftTop() {
		int dstTopRow = 9;
		int dstLeftCol = 6;
		int dstBottomRow = 11;
		int dstRightCol = 8;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		List<Ref> refs = new XRangeImpl(0, 0, sheet1, sheet1).removeIntersect(srcRef, dstRef);
		Ref q1 = refs.get(0);
		assertEquals(9, q1.getTopRow(), 1E-8);
		assertEquals(7, q1.getLeftCol(), 1E-8);
		assertEquals(9, q1.getBottomRow(), 1E-8);
		assertEquals(8, q1.getRightCol(), 1E-8);
		Ref q2 = refs.get(1);
		assertEquals(9, q2.getTopRow(), 1E-8);
		assertEquals(6, q2.getLeftCol(), 1E-8);
		assertEquals(9, q2.getBottomRow(), 1E-8);
		assertEquals(6, q2.getRightCol(), 1E-8);		
		Ref q3 = refs.get(2);
		assertEquals(10, q3.getTopRow(), 1E-8);
		assertEquals(6, q3.getLeftCol(), 1E-8);
		assertEquals(11, q3.getBottomRow(), 1E-8);
		assertEquals(6, q3.getRightCol(), 1E-8);
	}
	
	@Test
	public void testComplmentLeftTopReversely() {
		int dstTopRow = 9;
		int dstLeftCol = 6;
		int dstBottomRow = 11;
		int dstRightCol = 8;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		List<Ref> refs = new XRangeImpl(0, 0, sheet1, sheet1).removeIntersect(dstRef, srcRef);
		Ref q1 = refs.get(0);
		assertEquals(10, q1.getTopRow(), 1E-8);
		assertEquals(9, q1.getLeftCol(), 1E-8);
		assertEquals(11, q1.getBottomRow(), 1E-8);
		assertEquals(9, q1.getRightCol(), 1E-8);
		Ref q3 = refs.get(1);
		assertEquals(12, q3.getTopRow(), 1E-8);
		assertEquals(7, q3.getLeftCol(), 1E-8);
		assertEquals(12, q3.getBottomRow(), 1E-8);
		assertEquals(8, q3.getRightCol(), 1E-8);		
		Ref q4 = refs.get(2);
		assertEquals(12, q4.getTopRow(), 1E-8);
		assertEquals(9, q4.getLeftCol(), 1E-8);
		assertEquals(12, q4.getBottomRow(), 1E-8);
		assertEquals(9, q4.getRightCol(), 1E-8);
	}	
	
	@Test
	public void testOverlapRightTop() {
		int dstTopRow = 9;
		int dstLeftCol = 8;
		int dstBottomRow = 11;
		int dstRightCol = 10;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		int result = new XRangeImpl(0, 0, sheet1, sheet1).getOverlapState(srcRef, dstRef).ordinal();
		assertEquals(OverlapState.ON_RIGHT_TOP.ordinal(), result, 1E-8);
	}
	
	@Test
	public void testIntersectRightTop() {
		int dstTopRow = 9;
		int dstLeftCol = 8;
		int dstBottomRow = 11;
		int dstRightCol = 10;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		Ref resultRef = new XRangeImpl(0, 0, sheet1, sheet1).intersect(srcRef, dstRef);
		
		final int topRow = resultRef.getTopRow();
		final int bottomRow = resultRef.getBottomRow();
		final int leftCol = resultRef.getLeftCol();
		final int rightCol = resultRef.getRightCol();
		
		assertEquals(10, topRow, 1E-8);
		assertEquals(8, leftCol, 1E-8);
		assertEquals(11, bottomRow, 1E-8);
		assertEquals(9, rightCol, 1E-8);		
	}
	
	@Test
	public void testComplmentRightTop() {
		int dstTopRow = 9;
		int dstLeftCol = 8;
		int dstBottomRow = 11;
		int dstRightCol = 10;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		List<Ref> refs = new XRangeImpl(0, 0, sheet1, sheet1).removeIntersect(srcRef, dstRef);
		Ref q1 = refs.get(0);
		assertEquals(9, q1.getTopRow(), 1E-8);
		assertEquals(10, q1.getLeftCol(), 1E-8);
		assertEquals(9, q1.getBottomRow(), 1E-8);
		assertEquals(10, q1.getRightCol(), 1E-8);
		Ref q2 = refs.get(1);
		assertEquals(9, q2.getTopRow(), 1E-8);
		assertEquals(8, q2.getLeftCol(), 1E-8);
		assertEquals(9, q2.getBottomRow(), 1E-8);
		assertEquals(9, q2.getRightCol(), 1E-8);
		Ref q4 = refs.get(2);
		assertEquals(10, q4.getTopRow(), 1E-8);
		assertEquals(10, q4.getLeftCol(), 1E-8);
		assertEquals(11, q4.getBottomRow(), 1E-8);
		assertEquals(10, q4.getRightCol(), 1E-8);
	}	
	
	@Test
	public void testOverlapLeftBottom() {
		int dstTopRow = 11;
		int dstLeftCol = 6;
		int dstBottomRow = 13;
		int dstRightCol = 8;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		int result = new XRangeImpl(0, 0, sheet1, sheet1).getOverlapState(srcRef, dstRef).ordinal();
		assertEquals(OverlapState.ON_LEFT_BOTTOM.ordinal(), result, 1E-8);
	}
	
	@Test
	public void testIntersectLeftBottom() {
		int dstTopRow = 11;
		int dstLeftCol = 6;
		int dstBottomRow = 13;
		int dstRightCol = 8;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		Ref resultRef = new XRangeImpl(0, 0, sheet1, sheet1).intersect(srcRef, dstRef);
		
		final int topRow = resultRef.getTopRow();
		final int bottomRow = resultRef.getBottomRow();
		final int leftCol = resultRef.getLeftCol();
		final int rightCol = resultRef.getRightCol();
		
		assertEquals(11, topRow, 1E-8);
		assertEquals(7, leftCol, 1E-8);
		assertEquals(12, bottomRow, 1E-8);
		assertEquals(8, rightCol, 1E-8);		
	}
	
	@Test
	public void testComplmentLeftBottom() {
		int dstTopRow = 11;
		int dstLeftCol = 6;
		int dstBottomRow = 13;
		int dstRightCol = 8;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		List<Ref> refs = new XRangeImpl(0, 0, sheet1, sheet1).removeIntersect(srcRef, dstRef);
		Ref q2 = refs.get(0);
		assertEquals(11, q2.getTopRow(), 1E-8);
		assertEquals(6, q2.getLeftCol(), 1E-8);
		assertEquals(12, q2.getBottomRow(), 1E-8);
		assertEquals(6, q2.getRightCol(), 1E-8);
		Ref q3 = refs.get(1);
		assertEquals(13, q3.getTopRow(), 1E-8);
		assertEquals(6, q3.getLeftCol(), 1E-8);
		assertEquals(13, q3.getBottomRow(), 1E-8);
		assertEquals(6, q3.getRightCol(), 1E-8);		
		Ref q4 = refs.get(2);
		assertEquals(13, q4.getTopRow(), 1E-8);
		assertEquals(7, q4.getLeftCol(), 1E-8);
		assertEquals(13, q4.getBottomRow(), 1E-8);
		assertEquals(8, q4.getRightCol(), 1E-8);
	}	
	
	@Test
	public void testOverlapRightBottom() {
		int dstTopRow = 11;
		int dstLeftCol = 8;
		int dstBottomRow = 13;
		int dstRightCol = 10;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		int result = new XRangeImpl(0, 0, sheet1, sheet1).getOverlapState(srcRef, dstRef).ordinal();
		assertEquals(OverlapState.ON_RIGHT_BOTTOM.ordinal(), result, 1E-8);
	}
	
	@Test
	public void testIntersectRightBottom() {
		int dstTopRow = 11;
		int dstLeftCol = 8;
		int dstBottomRow = 13;
		int dstRightCol = 10;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		Ref resultRef = new XRangeImpl(0, 0, sheet1, sheet1).intersect(srcRef, dstRef);
		
		final int topRow = resultRef.getTopRow();
		final int bottomRow = resultRef.getBottomRow();
		final int leftCol = resultRef.getLeftCol();
		final int rightCol = resultRef.getRightCol();
		
		assertEquals(11, topRow, 1E-8);
		assertEquals(8, leftCol, 1E-8);
		assertEquals(12, bottomRow, 1E-8);
		assertEquals(9, rightCol, 1E-8);
	}
	
	@Test
	public void testComplmentRightBottom() {
		int dstTopRow = 11;
		int dstLeftCol = 8;
		int dstBottomRow = 13;
		int dstRightCol = 10;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		List<Ref> refs = new XRangeImpl(0, 0, sheet1, sheet1).removeIntersect(srcRef, dstRef);
		Ref q1 = refs.get(0);
		assertEquals(11, q1.getTopRow(), 1E-8);
		assertEquals(10, q1.getLeftCol(), 1E-8);
		assertEquals(12, q1.getBottomRow(), 1E-8);
		assertEquals(10, q1.getRightCol(), 1E-8);
		Ref q3 = refs.get(1);
		assertEquals(13, q3.getTopRow(), 1E-8);
		assertEquals(8, q3.getLeftCol(), 1E-8);
		assertEquals(13, q3.getBottomRow(), 1E-8);
		assertEquals(9, q3.getRightCol(), 1E-8);		
		Ref q4 = refs.get(2);
		assertEquals(13, q4.getTopRow(), 1E-8);
		assertEquals(10, q4.getLeftCol(), 1E-8);
		assertEquals(13, q4.getBottomRow(), 1E-8);
		assertEquals(10, q4.getRightCol(), 1E-8);
	}
	
	@Test
	public void testOverlapRight() {
		int dstTopRow = 10;
		int dstLeftCol = 9;
		int dstBottomRow = 12;
		int dstRightCol = 11;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		int result = new XRangeImpl(0, 0, sheet1, sheet1).getOverlapState(srcRef, dstRef).ordinal();
		assertEquals(OverlapState.ON_RIGHT.ordinal(), result, 1E-8);
	}	
	
	@Test
	public void testOverlapLeftWithSameRowCount() {
		int dstTopRow = 10;
		int dstLeftCol = 5;
		int dstBottomRow = 12;
		int dstRightCol = 7;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		int result = new XRangeImpl(0, 0, sheet1, sheet1).getOverlapState(srcRef, dstRef).ordinal();
		assertEquals(OverlapState.ON_LEFT.ordinal(), result, 1E-8);
	}
	
	@Test
	public void testOverlapTopWithSameColumnCount() {
		int dstTopRow = 8;
		int dstLeftCol = 7;
		int dstBottomRow = 10;
		int dstRightCol = 9;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		int result = new XRangeImpl(0, 0, sheet1, sheet1).getOverlapState(srcRef, dstRef).ordinal();
		assertEquals(OverlapState.ON_TOP.ordinal(), result, 1E-8);
	}
	
	@Test
	public void testOverlapSameRange() {
		int dstTopRow = 10;
		int dstLeftCol = 7;
		int dstBottomRow = 12;
		int dstRightCol = 9;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		int result = new XRangeImpl(0, 0, sheet1, sheet1).getOverlapState(srcRef, dstRef).ordinal();
		assertEquals(OverlapState.INCLUDED_INSIDE.ordinal(), result, 1E-8);
	}
	
//
//	@Test
//	public void testCopyToG10() {
//		
//		// 1 x 1, blank destination
//		int dstTopRow = 9;
//		int dstLeftCol = 6;
//		
//		// operation
//		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
//		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
//		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol);
//		srcRange.copy(dstRange); // copy src to dst
//		// validation
//		assertEquals(1, sheet1.getRow(dstTopRow).getCell(dstLeftCol).getNumericCellValue(), 1E-8);
//		assertEquals(2, sheet1.getRow(dstTopRow).getCell(dstLeftCol+1).getNumericCellValue(), 1E-8);
//		assertEquals(3, sheet1.getRow(dstTopRow).getCell(dstLeftCol+2).getNumericCellValue(), 1E-8);
//		// next row
//		assertEquals(4, sheet1.getRow(dstTopRow+1).getCell(dstLeftCol).getNumericCellValue(), 1E-8);
//		assertEquals(5, sheet1.getRow(dstTopRow+1).getCell(dstLeftCol+1).getNumericCellValue(), 1E-8);
//		assertEquals(6, sheet1.getRow(dstTopRow+1).getCell(dstLeftCol+2).getNumericCellValue(), 1E-8);
//		// next row
//		assertEquals(7, sheet1.getRow(dstTopRow+2).getCell(dstLeftCol).getNumericCellValue(), 1E-8);
//		assertEquals(8, sheet1.getRow(dstTopRow+2).getCell(dstLeftCol+1).getNumericCellValue(), 1E-8);
//		assertEquals(9, sheet1.getRow(dstTopRow+2).getCell(dstLeftCol+2).getNumericCellValue(), 1E-8);
//	}
//	
//	@Test
//	public void testCopyToG12() {
//		// 1 x 1, blank destination
//		int dstTopRow = 12;
//		int dstLeftCol = 6;
//		
//		// operation
//		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
//		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
//		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol);
//		srcRange.copy(dstRange); // copy src to dst
//		
//		// validation
//		assertEquals(1, sheet1.getRow(dstTopRow).getCell(dstLeftCol).getNumericCellValue(), 1E-8);
//		assertEquals(2, sheet1.getRow(dstTopRow).getCell(dstLeftCol+1).getNumericCellValue(), 1E-8);
//		assertEquals(3, sheet1.getRow(dstTopRow).getCell(dstLeftCol+2).getNumericCellValue(), 1E-8);
//		// next row
//		assertEquals(4, sheet1.getRow(dstTopRow+1).getCell(dstLeftCol).getNumericCellValue(), 1E-8);
//		assertEquals(5, sheet1.getRow(dstTopRow+1).getCell(dstLeftCol+1).getNumericCellValue(), 1E-8);
//		assertEquals(6, sheet1.getRow(dstTopRow+1).getCell(dstLeftCol+2).getNumericCellValue(), 1E-8);
//		// next row
//		assertEquals(7, sheet1.getRow(dstTopRow+2).getCell(dstLeftCol).getNumericCellValue(), 1E-8);
//		assertEquals(8, sheet1.getRow(dstTopRow+2).getCell(dstLeftCol+1).getNumericCellValue(), 1E-8);
//		assertEquals(9, sheet1.getRow(dstTopRow+2).getCell(dstLeftCol+2).getNumericCellValue(), 1E-8);		
//	}
//	
//	@Test 
//	public void testCopyToI12() {
//		int dstTopRow = 11;
//		int dstLeftCol = 8;
//		
//	}

}
