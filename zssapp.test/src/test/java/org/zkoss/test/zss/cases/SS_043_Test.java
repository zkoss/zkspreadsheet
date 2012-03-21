/* SS_043_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 20, 2012 4:28:54 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import junit.framework.Assert;

import org.junit.Test;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_043_Test extends ZSSAppTest {

	@Test
	public void number_format_dialog_number() {
		focus(0, 7);
		keyboardDirector.setEditText(0, 7, "12345.6789");
	
		mouseDirector.openCellContextMenu(0, 7);
		click(".zsmenuitem-formatCell");
		click("@listcell[label=\"Number\"] div.z-overflow-hidden");
		click("@listcell[label=\"1234.10\"] div.z-overflow-hidden");
		click("$_formatNumberDialog $okBtn");
		
		//verify
		Assert.assertEquals("12345.68", getCell(0, 7).getText());	
	}
	
	@Test
	public void number_format_dialog_currency() {
		focus(0, 10);
		keyboardDirector.setEditText(0, 10, "12345.6789");
	
		mouseDirector.openCellContextMenu(0, 10);
		click(".zsmenuitem-formatCell");
		click("@listcell[label=\"Currency\"] div.z-overflow-hidden");
		click("@window[title=\"Number Format\"] @listcell[label=\"($1,234.10)\"] div.z-overflow-hidden:eq(0)");
		click("$_formatNumberDialog $okBtn");
		
		//verify
		//Note. in Excel: after set this format, will add extra empty space
		Assert.assertEquals("$12,345.68 ", getCell(0, 10).getPureText().trim());	
	}
	
	@Test
	public void number_format_dialog_accounting() {
		focus(0, 10);
		keyboardDirector.setEditText(0, 10, "12345.6789");
	
		mouseDirector.openCellContextMenu(0, 10);
		click(".zsmenuitem-formatCell");
		click("@listcell[label=\"Accounting\"] div.z-overflow-hidden");
		click("@listcell[label=\"$1,234\"] div.z-overflow-hidden");
		click("$_formatNumberDialog $okBtn");
		
		//verify
		//Note. in Excel: after set this format, will add extra empty space
		Assert.assertEquals(" $12,345.68 ", getCell(0, 10).getPureText());
	}
	
	@Test
	public void number_format_dialog_date() {
		focus(0, 10);
		keyboardDirector.setEditText(0, 10, "12345.6789");
		
		mouseDirector.openCellContextMenu(0, 10);
		click(".zsmenuitem-formatCell");
		click("@listcell[label=\"Date\"] div.z-overflow-hidden");
		click("@listcell[label=\"3/14/01\"] div.z-overflow-hidden");
		click("$_formatNumberDialog $okBtn");
		
		Assert.assertEquals("10/18/33",  getCell(0, 10).getPureText());
	}
	
	@Test
	public void number_format_dialog_time() {
		focus(0, 10);
		keyboardDirector.setEditText(0, 10, "12345.6789");
	
		mouseDirector.openCellContextMenu(0, 10);
		click(".zsmenuitem-formatCell");
		click("@listcell[label=\"Time\"] div.z-overflow-hidden");
		click("@listcell[label=\"1:30 PM\"] div.z-overflow-hidden");
		click("$_formatNumberDialog $okBtn");
		
		//verify
		Assert.assertEquals("4:17 PM", getCell(0, 10).getPureText());
	}
	
	@Test
	public void number_format_dialog_percentage() {
		focus(0, 10);
		keyboardDirector.setEditText(0, 10, "12345.6789");
	
		mouseDirector.openCellContextMenu(0, 10);
		click(".zsmenuitem-formatCell");
		click("@listcell[label=\"Percentage\"] div.z-overflow-hidden");
		click("@listcell[label=\"percentage\"] div.z-overflow-hidden");
		click("$_formatNumberDialog $okBtn");
		
		//verify
		Assert.assertEquals("1234567.89%", getCell(0, 10).getPureText());
	}
	
	@Test
	public void number_format_dialog_fraction() {
		focus(0, 10);
		keyboardDirector.setEditText(0, 10, "12345.6789");
	
		mouseDirector.openCellContextMenu(0, 10);
		click(".zsmenuitem-formatCell");
		click("@listcell[label=\"Fraction\"] div.z-overflow-hidden");
		click("@listcell[label=\"Up to one digit(1/4)\"] div.z-overflow-hidden");
		click("$_formatNumberDialog $okBtn");
		
		//verify
		Assert.assertEquals("12345 2/3", getCell(0, 10).getPureText());
	}
	
	@Test
	public void number_format_dialog_scientific() {
		focus(0, 10);
		keyboardDirector.setEditText(0, 10, "12345.6789");
		
		mouseDirector.openCellContextMenu(0, 10);
		click(".zsmenuitem-formatCell");
		click("@listcell[label=\"Scientific\"] div.z-overflow-hidden");
		click("@listcell[label=\"scientific\"] div.z-overflow-hidden");
		click("$_formatNumberDialog $okBtn");
		
		//verify
		Assert.assertEquals("1.23E+04", getCell(0, 10).getPureText());
	}
	
	@Test
	public void number_format_dialog_text() {
		focus(0, 10);
		keyboardDirector.setEditText(0, 10, "12345.6789");
		
		mouseDirector.openCellContextMenu(0, 10);
		click(".zsmenuitem-formatCell");
		click("@listcell[label=\"Text\"] div.z-overflow-hidden");
		click("@listcell[label=\"text\"]");
		click("$_formatNumberDialog $okBtn");
		
		//verify
		Assert.assertEquals("12345.6789", getCell(0, 10).getText());	
	}
	
	@Test
	public void number_format_dialog_special() {
		focus(0, 10);
		keyboardDirector.setEditText(0, 10, "12345.6789");
		
		mouseDirector.openCellContextMenu(0, 10);
		click(".zsmenuitem-formatCell");
		click("@listcell[label=\"Text\"] div.z-overflow-hidden");
		click("@listcell[label=\"text\"]");
		click("$_formatNumberDialog $okBtn");
		
		//verify
		Assert.assertEquals("12345.6789", getCell(0, 10).getPureText());
	}
}
