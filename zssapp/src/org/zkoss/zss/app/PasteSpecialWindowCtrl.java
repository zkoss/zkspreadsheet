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

import java.util.HashMap;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.Utils;
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
		initSelectionMapper();
	}
	
	/**
	 * format selection i-18 String with Range ID
	 */
	private void initSelectionMapper() {
		pasteMapper.put(all, Range.PASTE_ALL);
		pasteMapper.put(allExcpetBorder, Range.PASTE_ALL_EXCEPT_BORDERS);
		pasteMapper.put(colWidth, Range.PASTE_COLUMN_WIDTHS);
		pasteMapper.put(comment, Range.PASTE_COMMENTS);
		pasteMapper.put(formula, Range.PASTE_FORMULAS);
		pasteMapper.put(formulaWithNum, Range.PASTE_FORMULAS_AND_NUMBER_FORMATS);
		pasteMapper.put(value, Range.PASTE_VALUES);
		pasteMapper.put(valueWithNumFmt, Range.PASTE_VALUES_AND_NUMBER_FORMATS);
		pasteMapper.put(fmt, Range.PASTE_FORMATS);
		pasteMapper.put(validation, Range.PASTE_VALIDATAION);
		
		opMapper.put(opAdd, Range.PASTEOP_ADD);
		opMapper.put(opSub, Range.PASTEOP_SUB);
		opMapper.put(opMul, Range.PASTEOP_MUL);
		opMapper.put(opDiv, Range.PASTEOP_DIV);
		opMapper.put(opNone, Range.PASTEOP_NONE);
	}
	
	private static Object getParam (String key) {
		return Executions.getCurrent().getArg().get(key);
	}
	

	
	public void onClick$okBtn() {
		okBtn.setDisabled(true);
		Utils.pasteSpecial(ss.getSelectedSheet(), ss.getHighlight(), 
				ss.getSelection().getTop(), ss.getSelection().getLeft(), 
				getPasteType(), getPasteOperation(),
				skipBlanks.isChecked(), transpose.isChecked());
		ss.setHighlight(null);
		pasteSpecialDlg.detach();
	}
	
	private int getPasteType() {
		Integer type = pasteMapper.get(pasteSelector.getSelectedItem());
		if (type == null)
			throw new NullPointerException("Paste type is null");
		return type;
	}
	
	private int getPasteOperation() {
		Integer op = opMapper.get(operationSelector.getSelectedItem());
		if (op == null)
			throw new NullPointerException("Paste operation is null");
		return op;
	}
}