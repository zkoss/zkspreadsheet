/* SS_238_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 19, 2012 3:13:39 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import junit.framework.Assert;

import org.junit.Test;
import org.zkoss.test.zss.Header;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_238_Test extends ZSSAppTest {
	
	/**
	 * drag column to hide it, then unhide
	 */
	@Test
	public void drag_column_hide_unhide () {
		final int COLUMN_F = 5;
		spreadsheet.focus(0, 5);
		Header header = spreadsheet.getTopPanel().getColumnHeader(COLUMN_F);
		Assert.assertTrue("Header shall be visible and has width", header.isVisible() && header.getWidth() > 0);
		
		mouseDirector.dragColumnToHide(COLUMN_F);
		Assert.assertTrue("Header shall be hidden", !header.isVisible());
		
		mouseDirector.dragColumnToResize(COLUMN_F, 50);
		Assert.assertTrue("Header shall be unhide", header.isVisible() && header.getWidth() == 50);
	}
	
	/**
	 * drag row to hide it, then unhide
	 */
	@Test
	public void drag_row_hide_unhide () {
		final int ROW_6 = 5;
		spreadsheet.focus(5, 0);
		
		Header header = spreadsheet.getLeftPanel().getRowHeader(ROW_6);
		Assert.assertTrue("Header shall be visible and has width", header.isVisible() && header.getHeight() > 0);
		
		mouseDirector.dragRowToHide(ROW_6);
		Assert.assertTrue("Header shall be hidden", !header.isVisible());
		
		mouseDirector.dragRowToResize(ROW_6, 50);
		Assert.assertTrue("Header shall be unhide", header.isVisible() && header.getHeight() == 50);
		
	}
}
