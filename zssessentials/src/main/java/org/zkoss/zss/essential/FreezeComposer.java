package org.zkoss.zss.essential;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Intbox;

/**
 * Demonstrate freeze API usage
 * 
 * @author Hawk
 * 
 */

@SuppressWarnings("serial")
public class FreezeComposer extends SelectorComposer<Component> {

	@Wire
	private Spreadsheet ss;

	@Listen("onClick = #freezeButton")
	public void freeze() {
		Ranges.range(ss.getSelectedSheet())
		.setFreezePanel(ss.getSelection().getRow(), ss.getSelection().getColumn());
	}
	
	@Listen("onClick = #unfreezeButton")
	public void unfreeze() {
		Ranges.range(ss.getSelectedSheet()).setFreezePanel(0,0);
	}
}
