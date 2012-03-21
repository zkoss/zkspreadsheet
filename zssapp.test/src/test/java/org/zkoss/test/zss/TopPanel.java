/* TopPanel.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 19, 2012 3:35:23 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss;

import org.openqa.selenium.WebDriver;
import org.zkoss.test.ConditionalTimeBlocker;
import org.zkoss.test.JQueryFactory;
import org.zkoss.test.Widget;

import com.google.inject.Inject;

/**
 * @author sam
 *
 */
public class TopPanel extends Widget {

	private final CellBlock.Factory cellBlockFactory;

	/**
	 * @param widgetScript
	 * @param webDriver
	 */
	@Inject
	/*package*/ TopPanel(SheetCtrl sheet, 
			CellBlock.Factory cellBlockFactory, JQueryFactory jqFactory, ConditionalTimeBlocker timeBlocker, WebDriver webDriver) {
		super(sheet.widgetScript() + ".tp" , jqFactory, timeBlocker, webDriver);
		
		this.cellBlockFactory = cellBlockFactory;
	}

	public Header getColumnHeader(int col) {
		String script = widgetScript() + ".getHeader(" + col + ")";
		return new Header(script, jqFactory, timeBlocker, webDriver);
	}
	
	public Integer getRowfreeze() {
		if (!hasCellBlock())
			return 0;
		
		return getCellBlock().getRowSize();
	}
	
	public boolean isRowfreeze() {
		if (!hasCellBlock())
			return false;
		
		CellBlock block = getCellBlock();
		return block.getRowSize() > 0;
	}
	
	private boolean hasCellBlock() {
		String hasBlockScript = "return " + widgetScript() + ".block != null";
		Boolean hasBlock = (Boolean) javascriptExecutor.executeScript(hasBlockScript);
		if (hasBlock == null) {
			return false;
		}
		return hasBlock;
	}
	
	public CellBlock getCellBlock() {
		CellBlock block = cellBlockFactory.create(widgetScript() + ".block");
		return block;
	}
}
