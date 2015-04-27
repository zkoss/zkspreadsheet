package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.*;

import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.Keys;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.EditorWidget;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;


public class Issue900Test extends ZSSTestCase {
	
	@Test
	public void testZSS901() throws Exception {
		getTo("issue3/901-filter-range.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		EditorWidget editor = ctrl.getInlineEditor();
		
		assertTrue(isDropdownOnCell(0, 1));
		assertTrue(isDropdownOnCell(1, 2));
		assertTrue(isDropdownOnCell(2, 3));
		
		
	}
	
	private boolean isDropdownOnCell(int dropdownIndex, int cellIndex) {
		JQuery dropdown = jq(".zsdropdown:eq(" + dropdownIndex + ")");
		int cellHorizontalMiddle = dropdown.offsetLeft() + dropdown.width()/2;
		int cellVerticalMiddle = dropdown.offsetTop() + dropdown.height()/2;

		int cellLeft = jq("@Row:eq(1) @Cell:eq(" + cellIndex + ")").offsetLeft();
		int cellRight = cellLeft + jq("@Row:eq(1) @Cell:eq(" + cellIndex + ")").width();
		int cellTop = jq("@Row:eq(1) @Cell:eq(" + cellIndex + ")").offsetTop();
		int cellBottom = cellTop + jq("@Row:eq(1) @Cell:eq(" + cellIndex + ")").height();
		
		return cellHorizontalMiddle > cellLeft && cellHorizontalMiddle < cellRight && cellVerticalMiddle > cellTop && cellVerticalMiddle < cellBottom;
	}
	
//	private boolean isCellHasBorder(JQuery cell, )
	
	@Ignore("vision, email")
	@Test
	public void testZSS900() throws Exception {}
	
	@Test
	public void testZSS903() throws Exception {
		getTo("issue3/903-paste-excel.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		EditorWidget editor = ctrl.getInlineEditor();
		
		click(ctrl.getCell("A1"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		editor.toWebElement().sendKeys("A1");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(ctrl.getCell("B1"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		editor.toWebElement().sendKeys("B1");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(ctrl.getCell("C1"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		editor.toWebElement().sendKeys("C1");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(ctrl.getCell("A2"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		editor.toWebElement().sendKeys("A2");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(ctrl.getCell("B2"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		editor.toWebElement().sendKeys("B2");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(ctrl.getCell("C2"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		editor.toWebElement().sendKeys("C2");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(ctrl.getCell("A1"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		dragAndDrop(ctrl.getCell("A1").$n(), ctrl.getCell("C2").$n());
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		jq("body").toWebElement().sendKeys(Keys.chord(Keys.CONTROL, "c"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		jq("$textbox").toWebElement().sendKeys(Keys.chord(Keys.CONTROL, "v"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(jq("$textbox").val(), "A1\tB1\tC1\nA2\tB2\tC2\n");
	}
}





