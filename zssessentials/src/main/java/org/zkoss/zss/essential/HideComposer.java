package org.zkoss.zss.essential;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.*;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * Demonstrate freeze API usage
 * 
 * @author Hawk
 * 
 */

@SuppressWarnings("serial")
public class HideComposer extends SelectorComposer<Component> {

	@Wire
	private Spreadsheet ss;

	@Listen("onClick = #hideButton")
	public void hideRow() {
		Range rowRange = Ranges.range(ss.getSelectedSheet(), ss.getSelection()).toRowRange();
		rowRange.setHidden(true);
		
//		CellOperationUtil.hide(rowRange);
//		CellOperationUtil.unhide(rowRange);
	}
	
	@Listen("onClick = #unhideButton")
	public void unhideRow() {
		Range rowRange = Ranges.range(ss.getSelectedSheet(), ss.getSelection()).toRowRange();
		rowRange.setHidden(false);
	}
}
