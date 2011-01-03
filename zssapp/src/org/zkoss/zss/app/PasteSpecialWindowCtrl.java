/* PasteSpecialWindowCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 1, 2010 6:05:07 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app;

import static org.zkoss.zss.app.base.Preconditions.checkNotNull;
import static org.zkoss.zss.app.cell.EditHelper.getPasteOperation;
import static org.zkoss.zss.app.cell.EditHelper.getPasteType;
import static org.zkoss.zss.app.cell.EditHelper.onPasteSpecial;

import java.util.HashMap;

import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.app.zul.Zssapps;
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
	private Dialog _pasteSpecialDialog;
	
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
		ss = checkNotNull(Zssapps.getSpreadsheetFromArg(), "Spreadsheet is null");
		if (ss.getHighlight() == null)
			throw new UiException("Spreadsheet must has highlight area as paste source, please set spreadsheet's highlight area");
	}
	
	private void init() {
		pasteSelector.setSelectedItem(all);
		operationSelector.setSelectedItem(opNone);
		
		skipBlanks.setChecked(false);
		transpose.setChecked(false);
		okBtn.setDisabled(false);
	}
	
	public void onOpen$_pasteSpecialDialog() {
		init();
		try {
			_pasteSpecialDialog.setMode(Window.MODAL);
		} catch (InterruptedException e) {
		}
	}

	public void onClick$okBtn() {
		okBtn.setDisabled(true);
		onPasteSpecial(ss, 
				getPasteType(pasteSelector.getSelectedItem().getValue()),
				getPasteOperation(operationSelector.getSelectedItem().getValue()),
				skipBlanks.isChecked(),
				transpose.isChecked());
		
		_pasteSpecialDialog.fireOnClose(null);
	}
}