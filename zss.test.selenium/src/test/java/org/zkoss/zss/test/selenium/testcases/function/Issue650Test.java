package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.fail;

import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.EditorWidget;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;

public class Issue650Test extends ZSSTestCase {
	
	@Test
	public void testZSS651() throws Exception {
		getTo("issue3/651-partial.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		
		click(ctrl.getCell("E3").toWebElement());
		jq("body").toWebElement().sendKeys(" 100");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		jq("body").toWebElement().sendKeys(Keys.ENTER);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(ctrl.getCell("G11").toWebElement().getText().trim(), "100");
		assertEquals(ctrl.getCell("G12").toWebElement().getText().trim(), "110");
		assertEquals(sheetChartUtil().getValue(".z-charts:eq(0)", 0, 4, "y"), "100");
		
		sheetFunction().gotoTab(2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(ctrl.getCell("G11").toWebElement().getText().trim(), "100");
		assertEquals(ctrl.getCell("G12").toWebElement().getText().trim(), "110");
		assertEquals(sheetChartUtil().getValue(".z-charts:eq(0)", 0, 4, "y"), "100");
		
		sheetFunction().gotoTab(1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
				
		click(ctrl.getCell("E3").toWebElement());
		jq("body").toWebElement().sendKeys(" 200");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		jq("body").toWebElement().sendKeys(Keys.ENTER);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(ctrl.getCell("G11").toWebElement().getText().trim(), "200");
		assertEquals(ctrl.getCell("G12").toWebElement().getText().trim(), "210");
		assertEquals(sheetChartUtil().getValue(".z-charts:eq(0)", 0, 4, "y"), 
				"200");
		
		sheetFunction().gotoTab(2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(ctrl.getCell("G11").toWebElement().getText().trim(), "200");
		assertEquals(ctrl.getCell("G12").toWebElement().getText().trim(), "210");
		assertEquals(sheetChartUtil().getValue(".z-charts:eq(0)", 0, 4, "y"), 
				"200");
		
		// 3.Reload page, insert column B, check H10,11,12's formula, they should be shifted. switch to Sheet 2, G10,11,12's formula shoulde be shifted
		getTo("issue3/651-partial.zul");
		ss = focusSheet();
		ctrl = ss.getSheetCtrl();
		rightClick(jq(".zstopcell:eq(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-insertSheetColumn:visible"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(ctrl.getCell("H11").toWebElement().getText().trim(), "5");
		assertEquals(ctrl.getCell("H12").toWebElement().getText().trim(), "15");
		assertEquals(sheetChartUtil().getValue(".z-charts:eq(0)", 0, 5, "y"), "5");
		
		sheetFunction().gotoTab(2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(ctrl.getCell("G11").toWebElement().getText().trim(), "5");
		assertEquals(ctrl.getCell("G12").toWebElement().getText().trim(), "15");
		assertEquals(sheetChartUtil().getValue(".z-charts:eq(0)", 0, 5, "y"), "5");
		
		// 4.Reload page, insert column B, change F3 to 100, H10,11,12 and chart shoulde change, switch to Sheet2 , G11, 12, and chart should change too
		
		getTo("issue3/651-partial.zul");
		ss = focusSheet();
		ctrl = ss.getSheetCtrl();
		rightClick(jq(".zstopcell:eq(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-insertSheetColumn:visible"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(ctrl.getCell("F3").toWebElement());
		jq("body").toWebElement().sendKeys(" 100");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		jq("body").toWebElement().sendKeys(Keys.ENTER);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(ctrl.getCell("H11").toWebElement().getText().trim(), "100");
		assertEquals(ctrl.getCell("H12").toWebElement().getText().trim(), "110");
		assertEquals(sheetChartUtil().getValue(".z-charts:eq(0)", 0, 5, "y"), "100");
		
		sheetFunction().gotoTab(2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(ctrl.getCell("G11").toWebElement().getText().trim(), "100");
		assertEquals(ctrl.getCell("G12").toWebElement().getText().trim(), "110");
		assertEquals(sheetChartUtil().getValue(".z-charts:eq(0)", 0, 5, "y"), "100");
		
		// 5.Redo insert row 1 on step 3,4
		
		getTo("issue3/651-partial.zul");
		ss = focusSheet();
		ctrl = ss.getSheetCtrl();
		rightClick(jq(".zsleftcell:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-insertSheetRow:visible"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(ctrl.getCell("G12").toWebElement().getText().trim(), "5");
		assertEquals(ctrl.getCell("G13").toWebElement().getText().trim(), "15");
		assertEquals(sheetChartUtil().getValue(".z-charts:eq(0)", 0, 4, "y"), "5");
		
		sheetFunction().gotoTab(2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(ctrl.getCell("G11").toWebElement().getText().trim(), "5");
		assertEquals(ctrl.getCell("G12").toWebElement().getText().trim(), "15");
		assertEquals(sheetChartUtil().getValue(".z-charts:eq(0)", 0, 4, "y"), "5");
		
		
		
		getTo("issue3/651-partial.zul");
		ss = focusSheet();
		ctrl = ss.getSheetCtrl();
		rightClick(jq(".zsleftcell:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-insertSheetRow:visible"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(ctrl.getCell("E4").toWebElement());
		jq("body").toWebElement().sendKeys(" 100");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		jq("body").toWebElement().sendKeys(Keys.ENTER);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(ctrl.getCell("G12").toWebElement().getText().trim(), "100");
		assertEquals(ctrl.getCell("G13").toWebElement().getText().trim(), "110");
		assertEquals(sheetChartUtil().getValue(".z-charts:eq(0)", 0, 4, "y"), "100");
		
		sheetFunction().gotoTab(2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(ctrl.getCell("G11").toWebElement().getText().trim(), "100");
		assertEquals(ctrl.getCell("G12").toWebElement().getText().trim(), "110");
		assertEquals(sheetChartUtil().getValue(".z-charts:eq(0)", 0, 4, "y"), "100");
	}
}
