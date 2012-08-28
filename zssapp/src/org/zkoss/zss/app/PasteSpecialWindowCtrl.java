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

import java.util.HashMap;

import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.app.zul.Zssapps;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zss.ui.sys.ActionHandler.Clipboard;
import org.zkoss.zul.Button;
import org.zkoss.zul.Checkbox;
import org.zkoss.zul.Messagebox;
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
		if (ss.getHighlight() == null) {
			Messagebox.show("Spreadsheet must has highlight area as paste source, please set spreadsheet's highlight area");
		}
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
		_pasteSpecialDialog.setMode(Window.MODAL);
	}

	public void onClick$okBtn() {
		okBtn.setDisabled(true);
		
		Clipboard clipboard = ss.getActionHandler().getClipboard();
		
		if (clipboard != null) {
			final Worksheet srcSheet = clipboard.sourceSheet;
			final Rect srcRect = clipboard.sourceRect;
			final Rect dst = ss.getSelection();
			
			Range rng = Utils.pasteSpecial(srcSheet, 
					srcRect, 
					ss.getSelectedSheet(), 
					dst.getTop(),
					dst.getLeft(),
					dst.getBottom(),
					dst.getRight(),
					getPasteType((String)pasteSelector.getSelectedItem().getValue()), 
					getPasteOperation((String)operationSelector.getSelectedItem().getValue()), 
					skipBlanks.isChecked(), transpose.isChecked());
			
			
			if (clipboard.type == Clipboard.Type.CUT) {
				Ranges
				.range(srcSheet, srcRect.getTop(), srcRect.getLeft(), srcRect.getBottom(), srcRect.getRight())
				.clearContents();
				
				final CellStyle defaultStyle = clipboard.book.createCellStyle();
				Ranges
				.range(srcSheet, srcRect.getTop(), srcRect.getLeft(),srcRect.getBottom(), srcRect.getRight())
				.setStyle(defaultStyle);
				
				ss.getActionHandler().clearClipboard();
				ss.setHighlight(null);
			}
			
			if (rng != null) {
				ss.setSelection(new Rect(rng.getColumn(), rng.getRow(), 
						rng.getLastColumn(), rng.getLastRow()));	
			}
		}
		
		_pasteSpecialDialog.fireOnClose(null);
	}
	
	/**
	 * Returns the paste operation base on i3-label, if no match
	 * <p> Default: returns {@link #Range.PASTEOP_NONE}, if no match 
	 * @param operation
	 * @return
	 */
	private static int getPasteOperation(String operation) {
		if (operation == null || "none".equals(operation) )
			return Range.PASTEOP_NONE;
		if ( "add".equals(operation) ) {
			return Range.PASTEOP_ADD;
		} else if ( "sub".equals(operation) ) {
			return Range.PASTEOP_SUB;
		} else if ( "mul".equals(operation) ) {
			return Range.PASTEOP_MUL;
		} else if ( "divide".equals(operation) ) {
			return Range.PASTEOP_DIV;
		}
		return Range.PASTEOP_NONE;
	}
	
	/**
	 * Returns the paste type base on i3 label (I18N)
	 * <p> Default: returns {@link #Range.PASTE_ALL}, if no match
	 * @return
	 */
	private static int getPasteType(String type) {
		if (type == null 
				|| "paste".equals(type)
				|| "all".equals(type) )
			return Range.PASTE_ALL;
		
		if ( "allExcpetBorder".equals(type) ) {
			return Range.PASTE_ALL_EXCEPT_BORDERS;
		} else if ( "columnWidth".equals(type) ) {
			return Range.PASTE_COLUMN_WIDTHS;
		} else if ( "comment".equals(type) ) {
			return Range.PASTE_COMMENTS;
		} else if ( "formula".equals(type) ) {
			return Range.PASTE_FORMULAS;
		} else if ( "formulaWithNumFmt".equals(type) ) {
			return Range.PASTE_FORMULAS_AND_NUMBER_FORMATS;
		} else if ( "value".equals(type) ) {
			return Range.PASTE_VALUES;
		} else if ( "valueWithNumFmt".equals(type) ) {
			return Range.PASTE_VALUES_AND_NUMBER_FORMATS;
		} else if ( "format".equals(type) ) {
			return Range.PASTE_FORMATS;
		} else if ( "validation".equals(type) ) {
			return Range.PASTE_VALIDATAION;
		}
		
		return Range.PASTE_ALL;
	}
}