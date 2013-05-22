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

//import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.PasteOperation;
import org.zkoss.zss.api.Range.PasteType;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.app.MainWindowCtrl.MainActionHandler;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.app.zul.Zssapps;
//import org.zkoss.zss.model.sys.XRange;
//import org.zkoss.zss.model.sys.XRanges;
//import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.ui.DefaultUserActionHandler.Clipboard;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zssex.ui.DefaultExUserActionHandler;
//import org.zkoss.zss.ui.impl.Utils;
//import org.zkoss.zss.ui.sys.XActionHandler.Clipboard;
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
		
		Clipboard clipboard = ((MainActionHandler)ss.getUserActionHandler()).getClipboard();
		
		if (clipboard != null) {
			Book book = ss.getBook();
			Sheet srcSheet = book.getSheet(clipboard.sourceSheetName);
			Sheet destSheet = ss.getSelectedSheet();
			final Rect srcRect = clipboard.sourceRect;
			final Rect destRect = ss.getSelection();
			
			Range src = Ranges.range(srcSheet,srcRect);
			Range dest = Ranges.range(destSheet,destRect);
			src.pasteSpecial(dest, getPasteType(pasteSelector.getSelectedItem().getValue().toString()), 
					getPasteOperation(operationSelector.getSelectedItem().getValue().toString()), 
					skipBlanks.isChecked(), transpose.isChecked());
			
			
			if (clipboard.type == Clipboard.Type.CUT) {
				Ranges.range(srcSheet, srcRect).clearContents();
				Ranges.range(srcSheet, srcRect).clearStyles();
				
				((MainActionHandler)ss.getUserActionHandler()).clearClipboard();
				ss.setHighlight(null);
			}
			
			if (src != null) {
				ss.setSelection(srcRect);	
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
	private static PasteOperation getPasteOperation(String operation) {
		if (operation == null || "none".equals(operation) )
			return PasteOperation.NONE;
		if ( "add".equals(operation) ) {
			return PasteOperation.ADD;
		} else if ( "sub".equals(operation) ) {
			return PasteOperation.SUB;
		} else if ( "mul".equals(operation) ) {
			return PasteOperation.MUL;
		} else if ( "divide".equals(operation) ) {
			return PasteOperation.DIV;
		}
		return PasteOperation.NONE;
	}
	
	/**
	 * Returns the paste type base on i3 label (I18N)
	 * <p> Default: returns {@link #Range.PASTE_ALL}, if no match
	 * @return
	 */
	private static PasteType getPasteType(String type) {
		if (type == null 
				|| "paste".equals(type)
				|| "all".equals(type) )
			return PasteType.ALL;
		
		if ( "allExcpetBorder".equals(type) ) {
			return PasteType.ALL_EXCEPT_BORDERS;
		} else if ( "columnWidth".equals(type) ) {
			return PasteType.COLUMN_WIDTHS;
		} else if ( "comment".equals(type) ) {
			return PasteType.COMMENTS;
		} else if ( "formula".equals(type) ) {
			return PasteType.FORMULAS;
		} else if ( "formulaWithNumFmt".equals(type) ) {
			return PasteType.FORMULAS_AND_NUMBER_FORMATS;
		} else if ( "value".equals(type) ) {
			return PasteType.VALUES;
		} else if ( "valueWithNumFmt".equals(type) ) {
			return PasteType.VALUES_AND_NUMBER_FORMATS;
		} else if ( "format".equals(type) ) {
			return PasteType.FORMATS;
		} else if ( "validation".equals(type) ) {
			return PasteType.VALIDATAION;
		}
		
		return PasteType.ALL;
	}
}