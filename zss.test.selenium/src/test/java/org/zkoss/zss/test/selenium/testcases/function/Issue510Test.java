package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.WebElement;
import org.zkoss.zss.test.selenium.AssertUtil;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;
import org.zkoss.zss.test.selenium.entity.EditorWidget;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;

public class Issue510Test extends ZSSTestCase {
	
	@Test
	public void testZSS512() throws Exception {
		getTo("/issue3/512-clear-merge.zul");
		
		CellWidget cellA1 = getSpreadsheet().getSheetCtrl().getCell("A1");
		click(jq("@button:eq(1)"));
		waitForTime(Setup.getTimeoutL0());
		assertTrue("Merge is no working", cellA1.isMerged());
		click(jq("@button:eq(2)"));
		waitForTime(Setup.getTimeoutL0());
		click(jq("@button:eq(3)"));
		waitForTime(Setup.getTimeoutL0());
		assertFalse("A1 should not be merged after clear all", cellA1.isMerged());
	}
	
	@Test
	public void testZSS513() throws Exception {
		getTo("/issue3/513-edit-validation.zul");
		
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		SheetCtrlWidget sheetCtrl = spreadsheet.getSheetCtrl();
		spreadsheet.focus();
		doubleClick(sheetCtrl.getCell("C3"));
		waitForTime(Setup.getTimeoutL0());
		EditorWidget editor = sheetCtrl.getInlineEditor();
		sendKeys(editor, "E");
		click(sheetCtrl.getCell("A1"));
		waitForTime(Setup.getTimeoutL0());
		assertTrue("Should show the warning", jq(".z-messagebox-window").exists());
		waitForTime(Setup.getTimeoutL0());
		click(jq(".z-messagebox-window button:eq(1)"));
		waitForTime(Setup.getTimeoutL0());
		assertEquals("Inline editor should stay in C3", "C3", (String) editor.getProperty("cellRef"));
	}
	
	@Test
	public void testZSS515() throws Exception {
		getTo("/issue3/515-deleteFreezeRow.zul");
		JQuery btns = jq(".z-button");
		
		for(int i = 0; i < 10; i++) {
			click(btns.get(i));
			waitForTime(Setup.getTimeoutL0());
			AssertUtil.assertNoJAVAError();			
		}
	}
	
	@Test
	public void testZSS515XLS() throws Exception {
		getTo("/issue3/515-deleteFreezeRow-xls.zul");
		JQuery btns = jq(".z-button");
		
		for(int i = 0; i < 10; i++) {
			click(btns.get(i));
			waitForTime(Setup.getTimeoutL0());
			AssertUtil.assertNoJAVAError();			
		}
	}
	
	@Test
	public void testZSS516() throws Exception {
		getTo("/issue3/516-setBackgroundWhite.zul");
		
		click(jq("@button:eq(0)"));
		waitForTime(Setup.getTimeoutL0());
		SheetCtrlWidget sheetCtrl = getSpreadsheet().getSheetCtrl();
		for (int row = 3; row <= 6; ++row) {
			for (int column = 1; column <= 9; ++column) {
				WebElement cell = sheetCtrl.getCell(row, column).toWebElement();
				assertEquals("Cell's background color should be purple",
						"rgba(102, 0, 255, 1)", cell.getCssValue("background-color"));
			}
		}
		click(jq("@button:eq(1)"));
		waitForTime(Setup.getTimeoutL0());
		for (int row = 3; row <= 6; ++row) {
			for (int column = 1; column <= 9; ++column) {
				WebElement cell = sheetCtrl.getCell(row, column).toWebElement();
				assertEquals("Cell's background color should be white",
						"rgba(255, 255, 255, 1)", cell.getCssValue("background-color"));
			}
		}
	}
	
	@Test
	public void testZSS517_1() throws Exception {
		getTo("/issue3/517-backgroundColor.zul");
		
		click(jq("@button:eq(0)"));
		waitForTime(Setup.getTimeoutL0());
		WebElement cell = getSpreadsheet().getSheetCtrl().getCell("B5").toWebElement();
		assertEquals("B5's background color should be purple",
				"rgba(102, 0, 255, 1)", cell.getCssValue("background-color"));
		click(jq("@button:eq(1)"));
		waitForTime(Setup.getTimeoutL0());
		assertEquals("B5's background color should be red",
				"rgba(255, 0, 0, 1)", cell.getCssValue("background-color"));
		click(jq("@button:eq(2)"));
		waitForTime(Setup.getTimeoutL0());
		assertEquals("B5's border color should be purple",
				"rgba(102, 0, 255, 1)", cell.getCssValue("border-bottom-color"));
		assertEquals("B5's bordercolor should be purple",
				"rgba(102, 0, 255, 1)", cell.getCssValue("border-right-color"));
		click(jq("@button:eq(3)"));
		waitForTime(Setup.getTimeoutL0());
		assertEquals("B5's border color should be red",
				"rgba(255, 0, 0, 1)", cell.getCssValue("border-bottom-color"));
		assertEquals("B5's border color should be red",
				"rgba(255, 0, 0, 1)", cell.getCssValue("border-right-color"));
	}
	
	@Test
	public void testZSS517_2() throws Exception {
		getTo("/issue3/517-backgroundColor2.zul");
		
		WebElement realCell = getSpreadsheet().getSheetCtrl().getCell("B5").$n("real");	
		click(jq("@button:eq(0)"));
		waitForTime(Setup.getTimeoutL0());
		assertEquals("B5's font color should be purple",
				"rgba(102, 0, 255, 1)", realCell.getCssValue("color"));
		click(jq("@button:eq(1)"));
		waitForTime(Setup.getTimeoutL0());
		assertEquals("B5's font color should be red",
				"rgba(255, 0, 0, 1)", realCell.getCssValue("color"));
	}
	
	@Ignore("Bug is not fixed yet")
	@Test
	public void testZSS518() throws Exception {
	}
	
	@Test
	public void testZSS519() throws Exception {
		getTo("/issue3/519-blankBackground.zul");
		SheetCtrlWidget sheetCtrl = getSpreadsheet().getSheetCtrl();
		for (int i = 1; i < 20; i++) {
			assertEquals("rgba(217, 149, 148, 1)", sheetCtrl.getCell("C" + i).$n()
					.getCssValue("background-color"));
			assertEquals("rgba(217, 149, 148, 1)", sheetCtrl.getCell("D" + i).$n()
					.getCssValue("background-color"));
		}
	}
	
	@Test
	public void testZSS519_2() throws Exception {
		getTo("/issue3/519-blankBackground2.zul");
		
		SheetCtrlWidget sheetCtrl1 = getSpreadsheet("@spreadsheet:eq(0)").getSheetCtrl();
		assertEquals("rgba(255, 0, 0, 1)", sheetCtrl1.getCell("D6").$n()
				.getCssValue("background-color"));
		assertEquals("rgba(255, 0, 0, 1)", sheetCtrl1.getCell("G6").$n()
				.getCssValue("background-color"));
		assertEquals("rgba(0, 176, 80, 1)", sheetCtrl1.getCell("D11").$n()
				.getCssValue("background-color"));
		assertEquals("rgba(255, 0, 0, 1)", sheetCtrl1.getCell("G11").$n()
				.getCssValue("background-color"));
	}
}
