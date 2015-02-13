package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.util.ArrayList;
import java.util.List;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.AssertUtil;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;

public class Issue490Test extends ZSSTestCase {

	@Test
	public void testZSS490() throws Exception {
		getTo("/issue3/490-externalReference.zul");
		
		click("@button:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		AssertUtil.assertNoJSError();
	}
	
	@Test
	public void testZSS491() throws Exception {
		getTo("/issue3/491-changeColumnChart.zul");
		
		click("@button:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		AssertUtil.assertNoJSError();
	}
	
	@Test
	public void testZSS492() throws Exception {
		getTo("/issue3/492-formula-rename-sheet.zul");
		
		click("@button:eq(1)");
		waitForTime(Setup.getTimeoutL0());
		click("@button:eq(2)");
		waitForTime(Setup.getTimeoutL0());
		assertEquals("abc", getSpreadsheet().getSheetCtrl().getCell("B1").getText());
	}
	
	@Test
	public void testZSS493() throws Exception {
		getTo("/issue3/493-deleteLastSheet.zul");
		
		click("@button:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		AssertUtil.assertNoJSError();
	}
	
	@Test
	public void testZSS494() throws Exception {
		getTo("issue3/494-reorder-sheet-break-formula.zul ");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		
		List<List<String>> rowValues = new ArrayList<List<String>>(5);
		
		for(int i = 1; i <= 5; i++) {
			sheetFunction().gotoTab(i);
			waitUntilProcessEnd(Setup.getTimeoutL0());
			rowValues.add(getCellContents(ctrl));
		}
		
		click(".z-button:eq(1)");
		
		// sheet 1
		sheetFunction().gotoTab(1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		List<String> cellValues = getCellContents(ctrl);
		List<String> cellValuesOrg = rowValues.get(0);
		for(int i = 0; i < 10; i++) {
			assertEquals(cellValues.get(i), cellValuesOrg.get(i));
		}
		
		// sheet 2
		sheetFunction().gotoTab(3);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		cellValues = getCellContents(ctrl);
		cellValuesOrg = rowValues.get(1);
		for(int i = 0; i < 10; i++) {
			assertEquals(cellValues.get(i), cellValuesOrg.get(i));
		}
		
		// sheet 3
		sheetFunction().gotoTab(2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		cellValues = getCellContents(ctrl);
		cellValuesOrg = rowValues.get(2);
		for(int i = 0; i < 10; i++) {
			assertEquals(cellValues.get(i), cellValuesOrg.get(i));
		}
		
		// sheet 4
		sheetFunction().gotoTab(4);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		cellValues = getCellContents(ctrl);
		cellValuesOrg = rowValues.get(3);
		for(int i = 0; i < 10; i++) {
			assertEquals(cellValues.get(i), cellValuesOrg.get(i));
		}
		
		// sheet 5
		sheetFunction().gotoTab(5);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		cellValues = getCellContents(ctrl);
		cellValuesOrg = rowValues.get(4);
		for(int i = 0; i < 10; i++) {
			assertEquals(cellValues.get(i), cellValuesOrg.get(i));
		}
	}

	private List<String> getCellContents(SheetCtrlWidget ctrl) {
		List<String> cellValues = new ArrayList<String>(10);
		for(int j = 0; j < 5; j++) {
			click(ctrl.getCell(1, j));
			waitUntilProcessEnd(Setup.getTimeoutL0());
			cellValues.add(jq(".zsformulabar-editor:eq(0)").text());
		}
		for(int j = 0; j < 5; j++) {
			click(ctrl.getCell(2, j));
			waitUntilProcessEnd(Setup.getTimeoutL0());
			cellValues.add(jq(".zsformulabar-editor:eq(0)").text());
		}
		return cellValues;
	}
	
	@Test
	public void testZSS496() throws Exception {
		getTo("issue3/496-insertPicture.zul");
		
		sheetFunction().gotoTab(2);
		waitForTime(Setup.getTimeoutL0());
		
		click(".zstbtn-insertPicture");
		waitForTime(Setup.getTimeoutL0());
		click(".z-button:eq(3)");
		waitForTime(Setup.getTimeoutL0());
		
		sheetFunction().gotoTab(1);
		waitForTime(Setup.getTimeoutL0());
		sheetFunction().gotoTab(2);
		waitForTime(Setup.getTimeoutL0());
		
		click(".zstbtn-insertPicture");
		waitForTime(Setup.getTimeoutL0());
		
		int windowIdx = Integer.parseInt(jq(".z-window:last").fun("index").toString());
		JQuery masks = jq(".z-modal-mask");
		int length = masks.length();
		
		for(int i = 0; i < length; i++) {
			assertTrue(windowIdx > Integer.parseInt(
					jq(".z-modal-mask:eq(" + i + ")").fun("index").toString()));
		}
	}
	
	@Test
	public void testZSS499() throws Exception {
		getTo("issue3/499-moveChartData.zul");
		
		click(".z-button:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		
		click(".z-button:eq(1)");
		waitForTime(Setup.getTimeoutL0());
		
		click(".z-button:eq(2)");
		waitForTime(Setup.getTimeoutL0());
		
		AssertUtil.assertNoJAVAError();
	}
	
	@Test
	public void testZSS499XLS() throws Exception {
		getTo("issue3/499-moveChartData-xls.zul");
		
		click(".z-button:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		
		click(".z-button:eq(1)");
		waitForTime(Setup.getTimeoutL0());
		
		click(".z-button:eq(2)");
		waitForTime(Setup.getTimeoutL0());
		
		AssertUtil.assertNoJAVAError();
	}
}
