package org.zkoss.zss.model.sys.impl;

import static org.junit.Assert.assertEquals;

import java.io.InputStream;
import java.util.List;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.XRanges;
import org.zkoss.zss.model.sys.XSheet;

/**
 * Test cut API.
 * Test the method "intersect" & "removeIntersect" which used in cut API.
 * @author kuro
 */
public class IntersectTest {
	
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
	public void testRemoveIntersectLeftTop() {
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
	public void testRemoveIntersectLeftTopReversely() {
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
	public void testRemoveIntersectRightTop() {
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
	public void testRemoveIntersectLeftBottom() {
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
	public void testRemoveIntersectRightBottom() {
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

}
