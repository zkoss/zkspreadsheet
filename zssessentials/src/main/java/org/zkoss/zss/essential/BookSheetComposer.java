/* BookSheetComposer.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 18, 2010 5:42:27 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.essential;

import java.util.ArrayList;
import java.util.List;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zkplus.databind.BindingListModelList;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Combobox;

/**
 * Demonstrate how to change different sheet
 * @author Hawk
 *
 */
public class BookSheetComposer extends SelectorComposer<Component>{
	
	@Wire
	Combobox sheetBox;
	@Wire
	Spreadsheet spreadsheet;
	//override
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		List<String> sheetNames = new ArrayList<String>();
		int sheetSize = spreadsheet.getBook().getNumberOfSheets();
		for (int i = 0; i < sheetSize; i++){
			sheetNames.add(spreadsheet.getBook().getSheetAt(i).getSheetName());
		}
		
		BindingListModelList model = new BindingListModelList(sheetNames, true);
		sheetBox.setModel(model);
	}
	
	@Listen("onSelect = #sheetBox")
	public void selectSheet(Event event) {
		spreadsheet.setSelectedSheet(sheetBox.getText());
	}
}