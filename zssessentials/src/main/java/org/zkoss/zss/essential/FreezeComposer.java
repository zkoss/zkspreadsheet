package org.zkoss.zss.essential;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Intbox;

/**
 * Demonstrate CellStyle, CellOperationUtil API usage
 * 
 * @author Hawk
 * 
 */
public class FreezeComposer extends SelectorComposer<Component> {

	@Wire
	private Intbox fRowBox;
	@Wire
	private Intbox fColumnBox;
	@Wire
	private Spreadsheet ss;


	@Listen("onClick = #freezeButton")
	public void freeze() {
		ss.setRowfreeze(fRowBox.getValue());
		ss.setColumnfreeze(fColumnBox.getValue());
	}
	
	
	@Listen("onClick = #unfreezeButton")
	public void unfreeze() {
		ss.setRowfreeze(-1);
		ss.setColumnfreeze(-1);
	}

}
