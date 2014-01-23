/* HeaderPositionHelperTest.java

	Purpose:
		
	Description:
		
	History:
		Aug 13, 2010 7:28:58 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.ui.impl;

import java.util.ArrayList;
import java.util.List;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import static org.junit.Assert.assertEquals;

import org.zkoss.zss.ui.impl.HeaderPositionHelper.HeaderPositionInfo;

/**
 * @author henrichen
 *
 */
public class HeaderPositionHelperTest {
	HeaderPositionHelper _helper;
	@Before
	public void setUp() throws Exception {
		int custId = 0;
		final List<HeaderPositionInfo> infos = new ArrayList<HeaderPositionInfo>();
		infos.add(new HeaderPositionInfo(3, 30, custId++, false, true));
		infos.add(new HeaderPositionInfo(9, 20, custId++, false, true));
		infos.add(new HeaderPositionInfo(10, 60, custId++, false, true));
		int defaultSize = 40;
		_helper = new HeaderPositionHelper(defaultSize, infos);
	}
	@After
	public void tearDown() throws Exception {
		_helper = null;
	}
	@Test
	public void testGetStartPixel() {
		assertEquals(0, _helper.getStartPixel(0));// 0
		assertEquals(40, _helper.getStartPixel(1));// 40
		assertEquals(120, _helper.getStartPixel(3));// 120
		assertEquals(150, _helper.getStartPixel(4));// 150
		assertEquals(190, _helper.getStartPixel(5));// 190
		assertEquals(310, _helper.getStartPixel(8));// 310
		assertEquals(350, _helper.getStartPixel(9));// 350
		assertEquals(370, _helper.getStartPixel(10));// 370
		assertEquals(430, _helper.getStartPixel(11));// 430
	}
	
	@Test 
	public void testGetCellIndex() {
		assertEquals(0,_helper.getCellIndex(39));// 0
		assertEquals(1,_helper.getCellIndex(40));// 1
		assertEquals(1,_helper.getCellIndex(79));// 1
		assertEquals(2,_helper.getCellIndex(80));// 2
		assertEquals(2,_helper.getCellIndex(119));// 2
		assertEquals(3,_helper.getCellIndex(120));// 3
		assertEquals(3,_helper.getCellIndex(149));// 3
		assertEquals(4,_helper.getCellIndex(150));// 4
		assertEquals(4,_helper.getCellIndex(189));// 4
		assertEquals(6,_helper.getCellIndex(235));// 6
		assertEquals(9,_helper.getCellIndex(350));// 9
		assertEquals(9,_helper.getCellIndex(369));// 9
		assertEquals(10,_helper.getCellIndex(370));// 10
		assertEquals(10,_helper.getCellIndex(429));// 10
		assertEquals(11,_helper.getCellIndex(430));// 11
		assertEquals(12,_helper.getCellIndex(480));// 12
	}
}
