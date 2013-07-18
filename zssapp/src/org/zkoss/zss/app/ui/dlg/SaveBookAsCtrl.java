/* 
	Purpose:
		
	Description:
		
	History:
		2013/7/10, Created by dennis

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.ui.dlg;

import java.util.Map;

import org.zkoss.lang.Strings;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zul.Textbox;
import org.zkoss.zul.Window;

/**
 * 
 * @author dennis
 *
 */
public class SaveBookAsCtrl extends DlgCtrlBase{
	private static final long serialVersionUID = 1L;
	
	public final static String ARG_NAME = "name";
	
	public static final String ON_SAVE = "onSave";
	
	@Wire
	Textbox bookName;
	
	private final static String URI = "~./zssapp/dlg/saveBookAs.zul";
	
	public static void show(EventListener<DlgCallbackEvent> callback,String name) {
		Map arg = newArg(callback);
		
		arg.put(ARG_NAME, name);
		
		Window comp = (Window)Executions.createComponents(URI, null, arg);
		comp.doModal();
		return;
	}
	
	@Listen("onClick=#save; onOK=#saveAsDlg")
	public void onSave(){
		if(Strings.isBlank(bookName.getValue())){
			bookName.setErrorMessage("empty name is not allowed");
			return;
		}
		postCallback(ON_SAVE, newMap(newEntry(ARG_NAME, bookName.getValue())));
		detach();
	}
	
	@Listen("onClick=#cancel; onCancel=#saveAsDlg")
	public void onCancel(){
		detach();
	}
}
