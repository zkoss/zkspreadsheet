package org.zkoss.zss.test.selenium.testcases.function;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.EditorWidget;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;


public class Issue970Test extends ZSSTestCase {
	
	@Ignore("need test in zssapp")
	@Test
	public void testZSS970() throws Exception {
		getTo("/#test.xlsx");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		EditorWidget editor = ctrl.getInlineEditor();
		int col = 0;
		int row = 0;
		String text = "c"; 
		for (int i = 0; i < 50; i++) {
			for (col = 0 ; col < 30 ; col++){
				for (row = 0 ; row <30 ; row++){
					click(ctrl.getCell(row, col));
					Thread.sleep(10);
					editor.toWebElement().sendKeys((i%10) + "");
					Thread.sleep(10);
				}
			}
			text = String.valueOf((char)(text.charAt(0) + 1));
		}
	}
}





