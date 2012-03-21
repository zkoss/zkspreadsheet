/* SS_042_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 20, 2012 3:59:48 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import junit.framework.Assert;

import org.junit.Test;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_042_Test extends ZSSAppTest {

	@Test
	public void insert_formula_dialog_search() {
		spreadsheet.focus(0, 10); 
		
		click(".zsformulabar-formulabtn");
		WebElement inp = jq("$searchTextbox").getWebElement();
		inp.sendKeys("bin");
		inp.sendKeys(Keys.ENTER);
		timeBlocker.waitResponse();
		if (browser.isSafari() || browser.isGecko()) {
			timeBlocker.waitUntil(1);
		}

		//verify
		String func1 = jq(".z-listcell-cnt:eq(0)").text();
		Assert.assertEquals("BIN2DEC", func1);
		String func2 = jq(".z-listcell-cnt:eq(1)").text();
		Assert.assertEquals("BIN2HEX", func2);
		String func6 = jq(".z-listcell-cnt:eq(5)").text();
		Assert.assertEquals("DEC2BIN", func6);
	}
	
	@Test
	public void insert_formula_dialog_category() {
		spreadsheet.focus(0, 10);
		click(".zsformulabar-formulabtn");
		click("$categoryCombobox i.z-combobox-btn");
		click("@comboitem[label=\"Text\"] td.z-comboitem-text");
		
		//verify
		String func1 = jq(".z-listcell-cnt:eq(0)").text();
		Assert.assertEquals("ASC", func1);
		String func2 = jq(".z-listcell-cnt:eq(1)").text();
		Assert.assertEquals("BAHTTEXT", func2);
		String func6 = jq(".z-listcell-cnt:eq(5)").text();
		Assert.assertEquals("CONCATENATE", func6);	
	}
	
	@Test
	public void insert_formula_dialog_select_formula() {
		spreadsheet.focus(0, 10);
		
		click(".zsformulabar-formulabtn");
		click("@listcell[label=\"ACCRINT\"]");		
		
		//verify
		String funcDesc = jq(".z-vlayout-inner .z-label:eq(1)").text();
		String expect = "ACCRINT(issue, first_interest, settlement, rate, par, frequency, basis, calc_method)";
		Assert.assertEquals(expect, funcDesc);
	}
	
	@Test
	public void insert_formula_dialog_ok() {
		spreadsheet.focus(0, 10);
		
		click(".zsformulabar-formulabtn");
		click("@listcell[label=\"ACCRINT\"]");		
		click("$okBtn");
		
		Assert.assertTrue(isVisible("$_composeFormulaDialog"));	
	}
	
	@Test
	public void compose_formula_dialog_textbox() {
		spreadsheet.focus(0, 10);
		
		click(".zsformulabar-formulabtn");
		WebElement inp = jq("$searchTextbox").getWebElement();
		inp.sendKeys("sum");
		inp.sendKeys(Keys.ENTER);
		timeBlocker.waitResponse();
		
		click("@listcell[label=\"SUM\"]");
		click("$_insertFormulaDialog $okBtn");
		click("$composeFormulaTextbox");
		inp = jq("$composeFormulaTextbox").getWebElement();
		inp.sendKeys("f7,f8,f9");
		inp.sendKeys(Keys.ENTER);
		timeBlocker.waitResponse();
		
		click("$_composeFormulaDialog $okBtn");
		
		//verify
		Assert.assertEquals("161500", getCell(0, 10).getText());
	}
	
	@Test
	public void compose_formula_dialog_arguments() {
		spreadsheet.focus(0, 10);
		
		click(".zsformulabar-formulabtn");
		WebElement inp = jq("$searchTextbox").getWebElement();
		inp.sendKeys("sum");
		inp.sendKeys(Keys.ENTER);
		timeBlocker.waitResponse();
		if (browser.isSafari() || browser.isGecko()) {
			timeBlocker.waitUntil(1);
		}
		
		click("@listcell[label=\"SUM\"]");
		click("$_insertFormulaDialog $okBtn");
		click("$composeFormulaTextbox");
		inp = jq("@window[mode=\"overlapped\"][title=\"Function Arguments\"] @textbox:eq(1)").getWebElement();
		inp.sendKeys("f7");
		inp.sendKeys(Keys.TAB);
		timeBlocker.waitResponse();
		if (browser.isSafari() || browser.isGecko()) {
			timeBlocker.waitUntil(1);
		}
			
		String classname = jq("$_composeFormulaDialog @textbox:eq(2)").attr("class");
		Assert.assertTrue(classname.contains("focus"));	
	}
}
