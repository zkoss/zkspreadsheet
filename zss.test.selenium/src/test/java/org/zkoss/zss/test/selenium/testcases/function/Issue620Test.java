package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.AssertUtil;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;
import org.zkoss.zss.test.selenium.entity.EditorWidget;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;
import org.zkoss.zss.test.selenium.entity.ZSStyle;

public class Issue620Test extends ZSSTestCase {
	
	@Test
	public void testZSS621() throws Exception {
		getTo("/issue3/621-shared-formula.zul");
		
		click(jq(".z-button:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".z-button:eq(2)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".z-button:eq(3)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(jq(".z-button:eq(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".z-button:eq(2)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".z-button:eq(3)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		AssertUtil.assertNoJAVAError();
	}
	
	@Test
	public void testZSS622() throws Exception {
		getTo("/issue3/622-runtime.zul");
		
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		SheetCtrlWidget sheetCtrl = spreadsheet.getSheetCtrl();
		CellWidget cellE7 = sheetCtrl.getCell("E7");
		EditorWidget editor = sheetCtrl.getInlineEditor();
		spreadsheet.focus();
		click(cellE7);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sendKeys(editor, "AAA");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(ZSStyle.FONT_BOLD_BTN));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals("Font weight should be bold in E7", "bold", cellE7.$n("real").getCssValue("font-weight"));		
	}
	
	@Test
	public void testZSS623() throws Exception {
		getTo("issue3/623-rowColumnStyle.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		
		click(jq(".z-button:eq(0)"));
		click(jq(".z-button:eq(1)"));
		click(jq(".z-button:eq(2)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		checkBgColor(false);
		
		assertEquals(ctrl.getCell("D2").$n().getText(), "A");
		assertEquals(ctrl.getCell("E4").$n().getText(), "A");
		assertEquals(ctrl.getCell("F5").$n().getText(), "A");
		
		click(jq(".z-button:eq(5)"));
		click(jq(".z-button:eq(4)"));
		click(jq(".z-button:eq(3)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		checkBgColor(true);
		
		click(jq(".z-button:eq(6)"));
		click(jq(".z-button:eq(0)"));
		click(jq(".z-button:eq(1)"));
		click(jq(".z-button:eq(2)"));
		click(jq(".z-button:eq(7)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(ctrl.getCell("D2").$n().getText(), "");
		assertEquals(ctrl.getCell("E4").$n().getText(), "");
		assertEquals(ctrl.getCell("F5").$n().getText(), "");
	}
	
	private void checkBgColor(boolean isWhite) {
		assertEquals(jq("@Row:eq(2) @Cell:eq(0)").css("background-color"), isWhite ? "rgba(0, 0, 0, 0)" : "rgb(255, 0, 0)");
		assertEquals(jq("@Row:eq(3) @Cell:eq(2)").css("background-color"), isWhite ? "rgba(0, 0, 0, 0)" : "rgb(255, 0, 0)");
		assertEquals(jq("@Row:eq(2) @Cell:eq(6)").css("background-color"), isWhite ? "rgba(0, 0, 0, 0)" : "rgb(255, 0, 0)");
		
		assertEquals(jq("@Row:eq(0) @Cell:eq(3)").css("background-color"), isWhite ? "rgba(0, 0, 0, 0)" : "rgb(0, 0, 255)");
		assertEquals(jq("@Row:eq(5) @Cell:eq(5)").css("background-color"), isWhite ? "rgba(0, 0, 0, 0)" : "rgb(0, 0, 255)");
		
		assertEquals(jq("@Row:eq(2) @Cell:eq(3)").css("background-color"), isWhite ? "rgba(0, 0, 0, 0)" : "rgb(0, 255, 0)");
		assertEquals(jq("@Row:eq(4) @Cell:eq(5)").css("background-color"), isWhite ? "rgba(0, 0, 0, 0)" : "rgb(0, 255, 0)");
	}

	@Test
	public void testZSS624() throws Exception {
		getTo("issue3/624-copy-same.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		
		rightClick(ctrl.getCell("A1"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-copy:visible"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		rightClick(ctrl.getCell("A1"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-paste:visible"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
				
		dragAndDrop(ctrl.getCell("A2").$n(), ctrl.getCell("B2").$n());
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		rightClick(ctrl.getCell("A2"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-copy:visible"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		rightClick(ctrl.getCell("A2"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-paste:visible"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		dragAndDrop(ctrl.getCell("A3").$n(), ctrl.getCell("B3").$n());
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		rightClick(ctrl.getCell("A3"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-copy:visible"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		dragAndDrop(ctrl.getCell("A3").$n(), ctrl.getCell("D3").$n());
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		rightClick(ctrl.getCell("A3"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-paste:visible"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		AssertUtil.assertNoJAVAError();
	}
}
