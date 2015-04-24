package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.*;

import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.EditorWidget;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;


public class Issue910Test extends ZSSTestCase {
	
	@Test
	public void testZSS915() throws Exception {
		getTo("issue3/915-indent-text.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		EditorWidget editor = ctrl.getInlineEditor();
		
//		JQuery a1 = jq(".zsrow:eq(0) .zscell:eq(0)");
//		int width = a1.width();
//		click(a1);
//		waitUntilProcessEnd(Setup.getTimeoutL0());
//		click(jq("@button:eq(0)"));
//		waitUntilProcessEnd(Setup.getTimeoutL0());
//		assertNotEquals(width, a1.width());
		
		
	}
	
	@Test
	public void testZSS918() throws Exception {
		getTo("issue3/918-vertical-text.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		EditorWidget editor = ctrl.getInlineEditor();
		
		JQuery boundQuery = jq(".zsrow:eq(1) .zscell:eq(1)");
		JQuery innerQuery = jq(".zsrow:eq(1) .zscell:eq(1) .zsvtxt");
		int boundWidth = boundQuery.width();
		int boundHeight = boundQuery.height();
		int innerWidth = innerQuery.width();
		int boundMiddle = boundQuery.offsetLeft() + boundWidth / 2;
		int innerMiddle = innerQuery.offsetLeft() + innerWidth / 2;
		assertTrue(Math.abs(boundMiddle - innerMiddle) < 3); // tolerance: middle alignment
		assertTrue(boundWidth < boundHeight); // vertical
		
		
		boundQuery = jq(".zsrow:eq(1) .zscell:eq(3)");
		innerQuery = jq(".zsrow:eq(1) .zscell:eq(3) .zsvtxt");
		boundWidth = boundQuery.width();
		boundHeight = boundQuery.height();
		innerWidth = innerQuery.width();
		boundMiddle = boundQuery.offsetLeft() + boundWidth / 2;
		innerMiddle = innerQuery.offsetLeft() + innerWidth / 2;
		assertTrue(Math.abs(boundMiddle - innerMiddle) < 3); // tolerance: middle alignment
		assertTrue(boundWidth < boundHeight); // vertical
		
		
		boundQuery = jq(".zsrow:eq(1) .zscell:eq(5)");
		innerQuery = jq(".zsrow:eq(1) .zscell:eq(5) .zsvtxt:eq(1)");
		boundWidth = boundQuery.width();
		boundHeight = boundQuery.height();
		boundMiddle = boundQuery.offsetLeft() + boundWidth / 2;
		innerMiddle = innerQuery.offsetLeft();
		assertTrue(Math.abs(boundMiddle - innerMiddle) < 5); // tolerance: middle alignment
		assertTrue(boundWidth < boundHeight); // vertical
	}
	

}





