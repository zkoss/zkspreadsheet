/* 
	Purpose:
		
	Description:
		
	History:
		2013/7/10, Created by dennis

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.ui.dlg;

import org.zkoss.lang.Strings;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zul.Textbox;

/**
 * 
 * @author dennis
 *
 */
public class SaveBookAsCtrl extends DlgCtrlBase{
	private static final long serialVersionUID = 1L;
	@Wire
	Textbox bookName;
	
	@Listen("onClick=#save; onOK=#saveAsDlg")
	public void onSave(){
		if(Strings.isBlank(bookName.getValue())){
			bookName.setErrorMessage("empty name is not allowed");
			return;
		}
		postCallback(DlgEvts.ON_SAVE, newMap(newEntry("name", bookName.getValue())));
		detach();
	}
	
	@Listen("onClick=#cancel; onCancel=#saveAsDlg")
	public void onCancel(){
		detach();
	}
}
