/* RefSheetTest.java

	Purpose:
		
	Description:
		
	History:
		Mar 8, 2010 2:25:23 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.engine.impl;


import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertSame;
import static org.junit.Assert.assertTrue;

import java.util.HashSet;
import java.util.Iterator;
import java.util.Random;
import java.util.Set;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.zss.engine.Ref;

/**
 * Unit test for {@link RefSheetImpl}.
 * @author henrichen
 *
 */
public class RefSheetTest {
	private static final String BOOK = "book1";
	private static final String SHEET_INDEX = "sheet1";
	private RefBookImpl _book;
	private RefSheetImpl _sheet;
	private Random _random;
	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		_book = (RefBookImpl) new RefBookImpl(BOOK, 64*1024-1, 255);
		_sheet = (RefSheetImpl) _book.getOrCreateRefSheet(SHEET_INDEX);
		_random = new Random();
	}

	/**
	 * @throws java.lang.Exception
	 */
	@After
	public void tearDown() throws Exception {
		_book = null;
		_sheet = null;
		_random = null;
	}
	
	@Test
	public void testGetOrCreateRef() {
		myTestGetOrCreateRef(0, 0, 0, 0); //A1:A1
		myTestGetOrCreateRef(0, 0, 1, 1); //A1:B2
	}
	
	@Test
	public void testGetOrCreateRefRandom() {
		//randomly generate refs and check
		for(int j = 0; j < 10000; ++j) {
			int row1 = _random.nextInt(1024*1024); //1M
			int col1 = _random.nextInt(16*1024); //16K
			int row2 = _random.nextInt(1024*1024); //1M
			int col2 = _random.nextInt(16*1024); //16K
			int tRow = Math.min(row1, row2);
			int lCol = Math.min(col1, col2);
			int bRow = Math.max(row1, row2);
			int rCol = Math.max(col1, col2);
			myTestGetOrCreateRef(tRow, lCol, bRow, rCol);
		}
	}
	
	@Test
	public void testGetRef() {
		assertNull(_sheet.getRef(0, 0, 0, 0));
		myTestGetOrCreateRef(0, 0, 0, 0); //A1
		assertNotNull(_sheet.getRef(0, 0, 0, 0));
	}
	
	@Test
	public void testRemoveRef() {
		myTestRemoveRef(0, 0, 0, 0);
		myTestRemoveRef(0, 0, 1, 1);
	}
	
	@Test
	public void testRemoveRefRandom() {
		//randomly generate refs and check
		for(int j = 0; j < 10000; ++j) {
			int row1 = _random.nextInt(1024*1024); //1M
			int col1 = _random.nextInt(16*1024); //16K
			int row2 = _random.nextInt(1024*1024); //1M
			int col2 = _random.nextInt(16*1024); //16K
			int tRow = Math.min(row1, row2);
			int lCol = Math.min(col1, col2);
			int bRow = Math.max(row1, row2);
			int rCol = Math.max(col1, col2);
			myTestRemoveRef(tRow, lCol, bRow, rCol);
		}
	}
	
	@Test
	public void testRemoveRefRandom2() {
		Set<Ref> refs = new HashSet<Ref>(10000);
		//randomly generate refs
		for(int j = 0; j < 10000; ++j) {
			int row1 = _random.nextInt(1024*1024); //1M
			int col1 = _random.nextInt(16*1024); //16K
			int row2 = _random.nextInt(1024*1024); //1M
			int col2 = _random.nextInt(16*1024); //16K
			int tRow = Math.min(row1, row2);
			int lCol = Math.min(col1, col2);
			int bRow = Math.max(row1, row2);
			int rCol = Math.max(col1, col2);
			refs.add(_sheet.getOrCreateRef(tRow, lCol, bRow, rCol));
		}
		for(Ref ref: refs) {
			int tRow = ref.getTopRow();
			int lCol = ref.getLeftCol();
			int bRow = ref.getBottomRow();
			int rCol = ref.getRightCol();
			Ref refx  = _sheet.removeRef(tRow, lCol, bRow, rCol);
			assertSame(ref, refx);
		}
	}
	
	@Test
	public void testGetHitRefs() {
		myTestGetHitRefs(10, 10, 10, 10);
	}
	
	@Test
	public void testGetHitRefsRandom() {
		//randomly generate refs and check
		for(int j = 0; j < 1000; ++j) {
			int row1 = _random.nextInt(1024*1024); //1M
			int col1 = _random.nextInt(16*1024); //16K
			int row2 = _random.nextInt(1024*1024); //1M
			int col2 = _random.nextInt(16*1024); //16K
			int tRow = Math.min(row1, row2);
			int lCol = Math.min(col1, col2);
			int bRow = Math.max(row1, row2);
			int rCol = Math.max(col1, col2);
			myTestGetHitRefs(tRow, lCol, bRow, rCol);
		}
	}
	
	@Test
	public void testGetHitRefsRandom2() {
		Set<Ref> refs = new HashSet<Ref>(1000);
		//randomly generate refs
		for(int j = 0; j < 1000; ++j) {
			int row1 = _random.nextInt(1024*1024); //1M
			int col1 = _random.nextInt(16*1024); //16K
			int row2 = _random.nextInt(1024*1024); //1M
			int col2 = _random.nextInt(16*1024); //16K
			int tRow = Math.min(row1, row2);
			int lCol = Math.min(col1, col2);
			int bRow = Math.max(row1, row2);
			int rCol = Math.max(col1, col2);
			refs.add(_sheet.getOrCreateRef(tRow, lCol, bRow, rCol));
		}
		
		//rendomly generate cell point and test
		for (int j = 0; j < 10000; ++j) {
			int row = _random.nextInt(1024*1024); //1M
			int col = _random.nextInt(16*1024); //16K
			Set<Ref> hits = _sheet.getHitRefs(row, col);
			
			for(Ref ref : refs) {
				if (hits.contains(ref)) {
					assertTrue(ref.getTopRow() <= row
							&& ref.getLeftCol() <= col
							&& ref.getBottomRow() >= row
							&& ref.getRightCol() >= col);
				} else {
					assertTrue(
							ref.getTopRow() > row 
							|| ref.getLeftCol() > col
							|| ref.getBottomRow() < row
							|| ref.getRightCol() < col);
				}
			}
		}
	}
	
	@Test
	public void testAddDependency() {
		//D1, A1:C1
		_sheet.addDependency(0, 3, null, 0, 0, 0, 2); //D1, A1:C1

		Set<Ref> hits = _sheet.getHitRefs(0,0);
		assertEquals(1, hits.size());
		Ref hit = hits.iterator().next();
		assertSame(_sheet.getRef(0,0,0,2), hit);
		assertEquals(1, hit.getDependents().size());
		Ref src = _sheet.getRef(0,3,0,3);
		assertSame(src, hit.getDependents().iterator().next());
		
		_sheet.removeDependency(0, 3, null, 0, 0, 0, 2); //D1, A1:C1
		assertTrue(hit.getDependents().isEmpty());
		assertTrue(src.getPrecedents().isEmpty());
	}
	
	@Test
	public void testGetDirectDependents() {
		//D1, A1:C1
		_sheet.addDependency(0, 3, null, 0, 0, 0, 2); //D1, A1:C1

		Set<Ref> dependents = _sheet.getDirectDependents(0,0); //change A1
		assertEquals(1, dependents.size());
		final Ref dependent = dependents.iterator().next();
		assertEquals(0, dependent.getTopRow());
		assertEquals(3, dependent.getLeftCol());
		assertEquals(_sheet, dependent.getOwnerSheet());
	}
	
	@Test
	public void testGetDependents() {
		//D1, A1:C1
		_sheet.addDependency(0, 3, null, 0, 0, 0, 2); //D1, A1:C1

		Set<Ref> dependents = _sheet.getAllDependents(0,0); //change A1
		assertEquals(1, dependents.size()); //D1
		final Ref dependent = dependents.iterator().next();
		assertEquals(0, dependent.getTopRow());
		assertEquals(3, dependent.getLeftCol());
		assertEquals(_sheet, dependent.getOwnerSheet());
	}
	
	@Test
	public void testGetDependentsDirectCircular() {
		//D1, A1:D1
		_sheet.addDependency(0, 3, null, 0, 0, 0, 3); //D1, A1:D1

		Set<Ref> dependents = _sheet.getAllDependents(0,0); //change A1
		assertEquals(1, dependents.size()); //D1
		final Ref dependent = dependents.iterator().next();
		assertEquals(0, dependent.getTopRow());
		assertEquals(3, dependent.getLeftCol());
		assertEquals(_sheet, dependent.getOwnerSheet());
	}
	
	@Test
	public void testGetDependentsIndirectCircular() {
		//D1, A1:C1
		_sheet.addDependency(0, 3, null, 0, 0, 0, 2); //D1, A1:C1
		_sheet.addDependency(0, 2, null, 0, 3, 1, 4); //C1, D1:E2

		Set<Ref> dependents = _sheet.getAllDependents(0,0); //change A1
		assertEquals(2, dependents.size()); //C1, D1
		final Iterator<Ref> it = dependents.iterator();
		Ref dependent = it.next();
		assertEquals(0, dependent.getTopRow());
		final int col = dependent.getLeftCol(); //2 or 3 
		assertTrue(col == 3 || col == 2);
		assertEquals(_sheet, dependent.getOwnerSheet());
		
		Ref dependent2 = it.next();
		assertEquals(0, dependent2.getTopRow());
		assertEquals(col == 3 ? 2 : 3, dependent2.getLeftCol());
		assertEquals(_sheet, dependent2.getOwnerSheet());
		
	}
	
	private void myTestGetOrCreateRef(int tRow, int lCol, int bRow, int rCol) {
		final Ref ref1 = _sheet.getOrCreateRef(tRow, lCol, bRow, rCol);
		assertNotNull(ref1);
		assertEqualRefs(tRow, lCol, bRow, rCol, ref1);
		
		final Ref ref2 = _sheet.getOrCreateRef(tRow, lCol, bRow, rCol);
		assertSame(ref1, ref2);

		final Ref ref3 = _sheet.getRef(tRow, lCol, bRow, rCol);
		assertSame(ref1, ref3);
	}
	
	private void assertEqualRefs(int tRow, int lCol, int bRow, int rCol, Ref ref) {
		assertEquals(tRow, ref.getTopRow());
		assertEquals(lCol, ref.getLeftCol());
		assertEquals(bRow, ref.getBottomRow());
		assertEquals(rCol, ref.getRightCol());
	}
	
	private void myTestRemoveRef(int tRow, int lCol, int bRow, int rCol) {
		assertNotNull(_sheet.getOrCreateRef(tRow, lCol, bRow, rCol));
		assertNotNull(_sheet.getRef(tRow, lCol, bRow, rCol));
		assertNotNull(_sheet.removeRef(tRow, lCol, bRow, rCol)); 
		assertNull(_sheet.getRef(0, 0, 0, 0));
		assertNull(_sheet.removeRef(tRow, lCol, bRow, rCol));
	}
	
	private void myTestGetHitRefs(int tRow, int lCol, int bRow, int rCol) {
		assertNotNull(_sheet.getOrCreateRef(tRow, lCol, bRow, rCol));
		hitTest(tRow, lCol, bRow, rCol);
		assertNotNull(_sheet.removeRef(tRow, lCol, bRow, rCol));
		assertNull(_sheet.getRef(tRow, lCol, bRow, rCol));
	}
	
	private void hitTest(int tRow, int lCol, int bRow, int rCol) {
		int row1 = tRow - 1 - _random.nextInt(100);
		int row2 = tRow + (bRow == tRow ? 0 : _random.nextInt(bRow - tRow));
		int row3 = bRow - (bRow == tRow ? 0 : _random.nextInt(bRow - tRow));
		int row4 = bRow + 1 + _random.nextInt(100);

		int col1 = lCol - 1 - _random.nextInt(100);
		int col2 = lCol + (rCol == lCol ? 0 : _random.nextInt(rCol - lCol));
		int col3 = rCol - (rCol == lCol ? 0 : _random.nextInt(rCol - lCol));
		int col4 = rCol + 1 + _random.nextInt(100);
		
		Set<Ref> refs = null;
		//row1
		refs = _sheet.getHitRefs(row1, col1);
		assertTrue(refs.isEmpty());
		
		refs = _sheet.getHitRefs(row1, col2);
		assertTrue(refs.isEmpty());
		
		refs = _sheet.getHitRefs(row1, col3);
		assertTrue(refs.isEmpty());
		
		refs = _sheet.getHitRefs(row1, col4);
		assertTrue(refs.isEmpty());

		//row2
		refs = _sheet.getHitRefs(row2, col1);
		assertTrue(refs.isEmpty());
		
		refs = _sheet.getHitRefs(row2, col2);
		assertTrue(!refs.isEmpty());
		
		refs = _sheet.getHitRefs(row2, col3);
		assertTrue(!refs.isEmpty());
		
		refs = _sheet.getHitRefs(row2, col4);
		assertTrue(refs.isEmpty());
		
		//row3
		refs = _sheet.getHitRefs(row3, col1);
		assertTrue(refs.isEmpty());
		
		refs = _sheet.getHitRefs(row3, col2);
		assertTrue(!refs.isEmpty());
		
		refs = _sheet.getHitRefs(row3, col3);
		assertTrue(!refs.isEmpty());
		
		refs = _sheet.getHitRefs(row3, col4);
		assertTrue(refs.isEmpty());

		//row4
		refs = _sheet.getHitRefs(row4, col1);
		assertTrue(refs.isEmpty());
		
		refs = _sheet.getHitRefs(row4, col2);
		assertTrue(refs.isEmpty());
		
		refs = _sheet.getHitRefs(row4, col3);
		assertTrue(refs.isEmpty());
		
		refs = _sheet.getHitRefs(row4, col4);
		assertTrue(refs.isEmpty());
	}
}
