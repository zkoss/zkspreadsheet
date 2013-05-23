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

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.event.SelectEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.app.zul.Zssapps;
import org.zkoss.zss.ui.Rect;
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
	private Listbox selectedCategory;
	
	private Button okBtn;
	
	private Spreadsheet spreadsheet;
	private Rect selection;
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		//TODO: move to WorkbookCtrl
		spreadsheet = Zssapps.getSpreadsheetFromArg();
		
		openFormatList("mfn_general");
	}
	
	public void onSelect$mfn_category(SelectEvent event) {
		openFormatList((String)mfn_category.getSelectedItem().getValue());
	}
	
	public void onOpen$_formatNumberDialog(ForwardEvent evt) {
		selection = (Rect) evt.getOrigin().getData();
		_formatNumberDialog.setMode(Window.MODAL);
	}
	
	public void openFormatList(String listId) {
		String[] myList = {"mfn_general","mfn_number","mfn_currency","mfn_accounting","mfn_date","mfn_time","mfn_percentage","mfn_fraction","mfn_scientific","mfn_text","mfn_special"};
		for(int i = 0; i< myList.length; i++){
			Listbox lb = (Listbox) self.getFellow(myList[i]);
			if(lb != null){
				if(listId.equals(myList[i])){
					lb.setVisible(true);
					lb.setSelectedIndex(0);
					selectedCategory = lb;
				}else{
					lb.setVisible(false);
				}
			}				
		}
	}

	public void onClick$okBtn() {
		Listitem  seldItem = mfn_category.getSelectedItem();
		if (seldItem == null) {
			showSelectFormatDialog();
			return;
		}
		if (selectedCategory == null || selectedCategory == mfn_general) {
			showSelectFormatDialog();
			return;
		}
		Listitem selectedItem = selectedCategory.getSelectedItem();

		if (selectedItem != null) {
			String formatCodes = selectedItem.getValue().toString();
			if (selection.getBottom() >= spreadsheet.getMaxrows())
				selection.setBottom(spreadsheet.getMaxrows() - 1);
			if (selection.getRight() >= spreadsheet.getMaxcolumns())
				selection.setRight(spreadsheet.getMaxcolumns() - 1);
			
			CellOperationUtil.applyDataFormat(Ranges.range(spreadsheet.getSelectedSheet(),selection), formatCodes);
		
		} else {
			showSelectFormatDialog();
			return;
		}
		_formatNumberDialog.fireOnClose(null);
	}

	private void showSelectFormatDialog() {
		Messagebox.show("Please select a category");
	}
}