package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;

import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;

public class Issue480Test extends ZSSTestCase {
	@Test
	public void testZSS485() throws Exception {
		getTo("/issue3/485-insert-row-copy.zul");
		
		click("@button:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		SheetCtrlWidget sheetCtrl = getSpreadsheet().getSheetCtrl();
		assertEquals("1", sheetCtrl.getCell("A4").getText());
	    assertEquals("2", sheetCtrl.getCell("B4").getText());
	    assertEquals("3", sheetCtrl.getCell("C4").getText());
	    assertEquals("4", sheetCtrl.getCell("A5").getText());
	    assertEquals("5", sheetCtrl.getCell("B5").getText());
	    assertEquals("6", sheetCtrl.getCell("C5").getText());
	    assertEquals("7", sheetCtrl.getCell("A6").getText());
	    assertEquals("8", sheetCtrl.getCell("B6").getText());
	    assertEquals("9", sheetCtrl.getCell("C6").getText());
	    assertEquals("a", sheetCtrl.getCell("A7").getText());
	    assertEquals("b", sheetCtrl.getCell("B7").getText());
	    assertEquals("c", sheetCtrl.getCell("C7").getText());
	    assertEquals("d", sheetCtrl.getCell("A8").getText());
	    assertEquals("e", sheetCtrl.getCell("B8").getText());
	    assertEquals("f", sheetCtrl.getCell("C8").getText());
	    assertEquals("g", sheetCtrl.getCell("A9").getText());
	    assertEquals("h", sheetCtrl.getCell("B9").getText());
	    assertEquals("i", sheetCtrl.getCell("C9").getText());
	}
}
