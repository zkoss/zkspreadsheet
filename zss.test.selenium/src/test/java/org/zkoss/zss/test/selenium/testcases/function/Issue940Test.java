package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.*;

import java.util.Map;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.EditorWidget;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;


public class Issue940Test extends ZSSTestCase {
	
	@Test
	public void testZSS944() throws Exception {
		getTo("/issue3/944-90deg-text.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		EditorWidget editor = ctrl.getInlineEditor();
		
		
		assertPosition("@Row:eq(0) @Cell:eq(0) .zscelltxt-real", 1, 1);
		assertPosition("@Row:eq(0) @Cell:eq(1) .zscelltxt-real", 1, 2);
		assertPosition("@Row:eq(0) @Cell:eq(2) .zscelltxt-real", 1, 3);

		assertPosition("@Row:eq(1) @Cell:eq(0) .zscelltxt-real", 1, 1);
		assertPosition("@Row:eq(1) @Cell:eq(1) .zscelltxt-real", 1, 2);
		assertPosition("@Row:eq(1) @Cell:eq(2) .zscelltxt-real", 1, 3);
		
		assertPosition("@Row:eq(0) @Cell:eq(4) .zscelltxt-real", 2, 1);
		assertPosition("@Row:eq(0) @Cell:eq(5) .zscelltxt-real", 2, 2);
		assertPosition("@Row:eq(0) @Cell:eq(6) .zscelltxt-real", 2, 3);

		assertPosition("@Row:eq(1) @Cell:eq(4) .zscelltxt-real", 2, 1);
		assertPosition("@Row:eq(1) @Cell:eq(5) .zscelltxt-real", 2, 2);
		assertPosition("@Row:eq(1) @Cell:eq(6) .zscelltxt-real", 2, 3);
		
		assertPosition("@Row:eq(0) @Cell:eq(8) .zscelltxt-real", 3, 1);
		assertPosition("@Row:eq(0) @Cell:eq(9) .zscelltxt-real", 3, 2);
		assertPosition("@Row:eq(0) @Cell:eq(10) .zscelltxt-real", 3, 3);

		assertPosition("@Row:eq(1) @Cell:eq(8) .zscelltxt-real", 3, 1);
		assertPosition("@Row:eq(1) @Cell:eq(9) .zscelltxt-real", 3, 2);
		assertPosition("@Row:eq(1) @Cell:eq(10) .zscelltxt-real", 3, 3);
	}
	
	private void assertPosition(String query, int valign, int halign) {
		Map<?, ?> rect = (Map<?, ?>) eval("return jq('" + query + "')[0].getBoundingClientRect()");
		double height = ((Number) rect.get("height")).doubleValue();
		double width = ((Number) rect.get("width")).doubleValue();
		double top = ((Number) rect.get("top")).doubleValue();
		double bottom = ((Number) rect.get("bottom")).doubleValue();
		double left = ((Number) rect.get("left")).doubleValue();
		double right = ((Number) rect.get("right")).doubleValue();
		double horiCenter = left + width/2;
		double verCenter = top + height/2;
		
		Map<?, ?> cellOffset = (Map<?, ?>) eval("return jq('" + query + "').closest('.zscell').offset()");
		double cellLeft = ((Number) cellOffset.get("left")).doubleValue();
		double cellTop = ((Number) cellOffset.get("top")).doubleValue();
		double cellWidth = ((Number) eval("return jq('" + query + "').closest('.zscell').width()")).doubleValue();
		double cellHeight = ((Number) eval("return jq('" + query + "').closest('.zscell').height()")).doubleValue();
		double cellBottom = cellTop + cellHeight;
		double cellRight = cellLeft + cellWidth;
		double cellHoriCenter = cellLeft + cellWidth/2;
		double cellVerCenter = cellTop + cellHeight/2;
		
		assertTrue(height > width);
		
		switch (valign) {
			case 1:
				// top
				assertTrue(Math.abs(cellTop - top) <= 5);			
				break;
			case 2:
				// middle
				assertTrue(Math.abs(cellVerCenter - verCenter) <= 5);			
				break;
			case 3:
				// bottom
				assertTrue(Math.abs(cellBottom - bottom) <= 5);			
				break;	
		}
		
		switch (halign) {
			case 1:
				// left
				assertTrue(Math.abs(cellLeft - left) <= 5);			
				break;
			case 2:
				// middle
				assertTrue(Math.abs(cellHoriCenter - horiCenter) <= 5);			
				break;
			case 3:
				// right
				assertTrue(Math.abs(cellRight - right) <= 5);			
				break;	
		}
	}
}





