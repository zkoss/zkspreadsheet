package org.zkoss.zss.model.sys.impl;

import static org.junit.Assert.assertEquals;

import java.io.InputStream;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.XRanges;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.model.sys.impl.XRangeImpl.OverlapState;

/**
 * Test overlap state of every condition.
 * There are 11 overlap states should be test.
 * NOT_OVERLAP, 
 * ON_RIGHT_BOTTOM, ON_LEFT_BOTTOM, ON_LEFT_TOP, ON_RIGHT_TOP, 
 * INCLUDED_INSIDE, INCLUDED_OUTSIDE, 
 * ON_TOP, ON_BOTTOM, ON_RIGHT, ON_LEFT.
 * two case when not overlap: 1: at the same sheet 2: not the same sheet.
 * @author kuro
 *
 */
public class OverlapStateTest {
	private static XBook _workbook;
	
	// 3 x 3, 1-9 matrix (source)
	// Source Range (H11, J13)
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
	
	/**
	 * destination is any range in the same sheet with source range but doesn't overlap with each other.
	 */
	@Test
	public void testNotOverlapSameSheet() {
		
		// any range
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
	
	/**
	 * test the case that the "source range and destination range are in the different sheet".
	 * check there overlap state should be "not overlap".
	 */
	@Test
	public void testNotOverlapNotSameSheet() {
		
		// set the same coordinate as source (overlap) 
		int dstTopRow = srcTopRow;
		int dstLeftCol = srcLeftCol;
		int dstBottomRow = srcBottomRow;
		int dstRightCol = srcRightCol;
		
		// but initialize the range at different sheet
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XSheet dstSheet = _workbook.getWorksheet("Sheet2");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(dstSheet, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		// which should result in "not overlap"
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		int result = new XRangeImpl(0, 0, sheet1, sheet1).getOverlapState(srcRef, dstRef).ordinal();
		assertEquals(OverlapState.NOT_OVERLAP.ordinal(), result, 1E-8);
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
	
	/**
	 * destination range is right hand side of source range.
	 * Source Range (H11, J13).
	 * Destination Range (J11, L13).
	 */
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
	
	/**
	 * destination range is left hand side of source range.
	 * Source Range (H11, J13).
	 * Destination Range (F11, H13).
	 */
	@Test
	public void testOverlapLeft() {
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
	public void testOverlapTop() {
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
	
	/**
	 * check overlap state when the source ref is the same with destination ref (four corner are the same)
	 * it should return the overlap state "include_inside" 
	 */
	@Test
	public void testOverlapSameRange() {
		int dstTopRow = srcTopRow;
		int dstLeftCol = srcLeftCol;
		int dstBottomRow = srcBottomRow;
		int dstRightCol = srcRightCol;
		
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		XRange srcRange = XRanges.range(sheet1, srcTopRow, srcLeftCol, srcBottomRow, srcRightCol);
		XRange dstRange = XRanges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		
		Ref srcRef = ((XRangeImpl)srcRange).getRefs().iterator().next();
		Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
		int result = new XRangeImpl(0, 0, sheet1, sheet1).getOverlapState(srcRef, dstRef).ordinal();
		assertEquals(OverlapState.INCLUDED_INSIDE.ordinal(), result, 1E-8);
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

}
