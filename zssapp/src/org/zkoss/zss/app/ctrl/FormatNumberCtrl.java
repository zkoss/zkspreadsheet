/* FormatNumberCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Dec 1, 2010 11:32:52 AM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.ctrl;

import java.util.HashMap;
import java.util.Map;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.app.zul.Zssapps;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Button;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Listitem;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.Window;

/**
 * @author Sam
 *
 */
public class FormatNumberCtrl extends GenericForwardComposer {
	
	private Dialog _formatNumberDialog;
	private Listbox mfn_category;
	private Listbox mfn_general;
	
	private Button okBtn;
	
	private Spreadsheet spreadsheet;
	
	Map mapLabelListbox;
	public FormatNumberCtrl() {
		mapLabelListbox = new HashMap();

		mapLabelListbox.put("General", "mfn_general");
		mapLabelListbox.put("Number", "mfn_number");
		mapLabelListbox.put("Currency", "mfn_currency");
		mapLabelListbox.put("Accounting", "mfn_accounting");
		mapLabelListbox.put("Date", "mfn_date");
		mapLabelListbox.put("Time", "mfn_time");
		mapLabelListbox.put("Percentage", "mfn_percentage");
		mapLabelListbox.put("Fraction", "mfn_fraction");
		mapLabelListbox.put("Scientific", "mfn_scientific");
		mapLabelListbox.put("Text", "mfn_text");
		mapLabelListbox.put("Special", "mfn_special");
	}
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		spreadsheet = Zssapps.getSpreadsheetFromArg();
	}
	
	public void onOpen$_formatNumberDialog() {
		try {
			_formatNumberDialog.setMode(Window.MODAL);
		} catch (InterruptedException e) {
		}
	}

	public void onClick$okBtn() {
		Listitem  seldItem = mfn_category.getSelectedItem();
		if (seldItem == null) {
			try {
				Messagebox.show("Please select a category");
			} catch (InterruptedException e) {
			}
			return;
		}
		Listbox selectedList = (Listbox)self.getFellow((String)mapLabelListbox.get(seldItem.getLabel()));
		Listitem selectedItem = selectedList.getSelectedItem();

		if (selectedItem != null) {
			String formatCodes = selectedItem.getValue().toString();
			Utils.setDataFormat(spreadsheet.getSelectedSheet(), spreadsheet.getSelection(), formatCodes);			
		} else {
			try {
				Messagebox.show("Please select a format");
			} catch (InterruptedException e) {
			}
		}
		_formatNumberDialog.fireOnClose(null);
	}
}