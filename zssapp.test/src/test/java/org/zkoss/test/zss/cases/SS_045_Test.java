/* SS_045_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 20, 2012 5:40:59 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import org.junit.Assert;
import org.junit.Test;
import org.zkoss.test.zss.CellCache;
import org.zkoss.test.zss.CellCacheAggeration;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_045_Test extends ZSSAppTest {
	
	@Test
	public void drag_edit() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration cache = builder.build();
		mouseDirector.dragEdit(tRow, lCol, bRow, rCol, 12, 10);
		
		Assert.assertEquals(cache, builder.offset(12, 10).build());
	}
	
	@Test
	public void drag_fill() {
		CellCache J12 = getCellCache(11, 9);
		mouseDirector.dragFill(11, 9, 11, 11);
		
		Assert.assertEquals(J12.getText(), getCell(11, 10).getText());
		Assert.assertEquals(J12.getFillColor(), getCell(11, 10).getFillColor());
	}
	
	
	@Test
	public void auto_fill_month_increase() {
		keyboardDirector.setEditText(12, 10, 13, 10, new String[] { "Jan", "Feb" });
		mouseDirector.dragFill(12, 10, 13, 10, 16, 10);

		Assert.assertEquals("Mar", getCell(14, 10).getText());
		Assert.assertEquals("Apr", getCell(15, 10).getText());
		Assert.assertEquals("May", getCell(16, 10).getText());
	}
	
	@Test
	public void auto_fill_date_increase() {
		keyboardDirector.setEditText(12, 10, 13, 10, new String[] {"4/1", "4/2"});
		mouseDirector.dragFill(12, 10, 13, 10, 16, 10);
		
		Assert.assertEquals("3-Apr", getCell(14, 10).getText());
		Assert.assertEquals("4-Apr", getCell(15, 10).getText());
		Assert.assertEquals("5-Apr", getCell(16, 10).getText());
	}
	
	@Test
	public void auto_fill_week_increase() {
		keyboardDirector.setEditText(12, 10, 13, 10, new String[] {"Mon", "Tue"});
		mouseDirector.dragFill(12, 10, 13, 10, 16, 10);
		
		Assert.assertEquals("Wed", getCell(14, 10).getText());
		Assert.assertEquals("Thu", getCell(15, 10).getText());
		Assert.assertEquals("Fri", getCell(16, 10).getText());
	}
	
	@Test
	public void auto_fill_number_linear_increase() {
		keyboardDirector.setEditText(12, 10, 13, 10, new String[] {"1", "2"});
		mouseDirector.dragFill(12, 10, 13, 10, 16, 10);
		
		Assert.assertEquals("3", getCell(14, 10).getText());
		Assert.assertEquals("4", getCell(15, 10).getText());
		Assert.assertEquals("5", getCell(16, 10).getText());	
	}
	
	@Test
	public void auto_fill_number_multiply_increase() {
		keyboardDirector.setEditText(12, 10, 13, 10, new String[] {"1", "4"});
		mouseDirector.dragFill(12, 10, 13, 10, 16, 10);
		
		Assert.assertEquals("7", getCell(14, 10).getText());
		Assert.assertEquals("10", getCell(15, 10).getText());
		Assert.assertEquals("13", getCell(16, 10).getText());
	}
}
