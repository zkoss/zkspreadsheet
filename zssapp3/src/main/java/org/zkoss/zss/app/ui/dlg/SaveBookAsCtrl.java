package org.zkoss.zss.app.ui.dlg;

import org.zkoss.lang.Strings;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zul.Textbox;


public class SaveBookAsCtrl extends DlgCtrlBase{
	private static final long serialVersionUID = 1L;
	@Wire
	Textbox bookName;
	
	@Listen("onClick=#save")
	public void doSave(){
		if(Strings.isBlank(bookName.getValue())){
			bookName.setErrorMessage("empty name is not allowed");
			return;
		}
		postCallback(Dlgs.ON_SAVE, newMap(newEntry("name", bookName.getValue())));
		detach();
	}
	
	@Listen("onClick=#cancel")
	public void doCancel(){
		detach();
	}
}
