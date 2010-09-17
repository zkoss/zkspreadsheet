/* PasteSpecialWindowCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 1, 2010 6:05:07 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.app;

import static org.zkoss.zss.app.event.EditHelper.getPasteOperation;
import static org.zkoss.zss.app.event.EditHelper.getPasteType;
import static org.zkoss.zss.app.event.EditHelper.onPasteSpecial;

import java.util.HashMap;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Button;
import org.zkoss.zul.Checkbox;
import org.zkoss.zul.Radio;
import org.zkoss.zul.Radiogroup;
import org.zkoss.zul.Window;
/**
 * @author Sam
 *
 */
public class PasteSpecialWindowCtrl extends GenericForwardComposer {

	private Spreadsheet ss;
	private Window pasteSpecialDlg;
	
	private Radiogroup pasteSelector;
	private Radio all;
	private Radio allExcpetBorder;
	private Radio colWidth;
	private Radio comment;
	private Radio formula;
	private Radio formulaWithNum;
	private Radio value;
	private Radio valueWithNumFmt;
	private Radio fmt;
	private Radio validation;
	private HashMap<Radio, Integer> pasteMapper = new HashMap<Radio, Integer>(10);
	
	private Radiogroup operationSelector;
	private Radio opAdd;
	private Radio opSub;
	private Radio opMul;
	private Radio opDiv;
	private Radio opNone;
	private HashMap<Radio, Integer> opMapper = new HashMap<Radio, Integer>(5);
	
	private Button okBtn;
	private Checkbox skipBlanks;
	private Checkbox transpose;

	public PasteSpecialWindowCtrl () {
		ss = (Spreadsheet)getParam("spreadsheet");
		if (ss == null)
			throw new UiException("Spreadsheet object is empty");
		if (ss.getHighlight() == null)
			throw new UiException("Spreadsheet must has highlight area as paste source, please set spreadsheet's highlight area");
	}
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		pasteSelector.setSelectedItem(all);
		operationSelector.setSelectedItem(opNone);
	}
	
	private static Object getParam (String key) {
		return Executions.getCurrent().getArg().get(key);
	}

	public void onClick$okBtn() {
		okBtn.setDisabled(true);
		onPasteSpecial(ss, 
				getPasteType(pasteSelector.getSelectedItem().getLabel()),
				getPasteOperation(operationSelector.getSelectedItem().getLabel()),
				skipBlanks.isChecked(),
				transpose.isChecked());
		pasteSpecialDlg.detach();
	}
}