package org.zkoss.test.zss.cases;
import junit.framework.Assert;

import org.junit.Test;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

import com.google.common.base.Strings;

/* SS_047_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 18, 2012 9:39:28 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
 */

/**
 * 
 * @author sam
 *
 */
@ZSSTestCase
public class SS_047_Test extends ZSSAppTest {

	/**
	 * 1. select a cell (if not empty cell, remove value use Delete key)
	 * 2. press a key shall enter edit mode
	 * 3. press Enter shall change the cell value
	 */
	@Test
	public void inline_editing_enter() {		
		keyboardDirector.setEditText(12, 5, "a"); //setEditText press enter at end
		Assert.assertTrue("Cell text shall be 'a'", "a".equals(getCell(12, 5).getEdit()));
	}
	
	
	/**
	 * 1. get cell's original text
	 * 2. editing and press ESC to cancel editing
	 * 3. cell text shall not change
	 */
	@Test
	public void inline_editing_cancel() {
		
		String originalText = getCell(12, 5).getEdit();
		
		keyboardDirector.sendKeys(12, 5, "abcdefg");
		keyboardDirector.esc(spreadsheet.getInlineEditor().jq$n());
		
		String text = getCell(12, 5).getEdit();
		Assert.assertTrue("Cell text shall not changed", originalText.equals(text));
	}
}
