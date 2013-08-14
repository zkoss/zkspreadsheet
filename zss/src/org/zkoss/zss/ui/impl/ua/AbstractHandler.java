/* AbstractUserHandler.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/8/2 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ui.impl.ua;

import org.zkoss.util.resource.Labels;
import org.zkoss.zss.api.IllegalOpArgumentException;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Rect;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.UserActionContext;
import org.zkoss.zss.ui.UserActionHandler;
import org.zkoss.zul.Messagebox;

/**
 * @author dennis
 * @since 3.0.0
 */
public abstract class AbstractHandler implements UserActionHandler{

	@Override
	public boolean isEnabled(Book book, Sheet sheet) {
		return true;
	}
	
	
	protected void showProtectMessage() {
		String message = Labels.getLabel("zss.actionhandler.msg.sheet_protected");
		showWarnMessage(message);
	}
	
	protected void showInfoMessage(String message) {
		String title = Labels.getLabel("zss.actionhandler.msg.info_title");
		Messagebox.show(message, title, Messagebox.OK, Messagebox.INFORMATION);
	}
	
	protected void showWarnMessage(String message) {
		String title = Labels.getLabel("zss.actionhandler.msg.warn_title");
		Messagebox.show(message, title, Messagebox.OK, Messagebox.EXCLAMATION);
	}
	
	public boolean process(UserActionContext ctx){
		try{
			return processAction(ctx);
		}catch(IllegalOpArgumentException x){
			showInfoMessage(Labels.getLabel("zss.actionhandler.msg.illegal_range_operation")+" : "+x.getMessage());
			return true;
		}
	}

	protected boolean checkInFreezePanel(Range range) {
		Sheet sheet = range.getSheet();
		int fzr = sheet.getRowFreeze();//to index
		int fzc = sheet.getColumnFreeze();
		
		if(fzr>range.getRow() || fzc > range.getColumn()){
			return true;
		}
		return false;
	}
	protected abstract boolean processAction(UserActionContext ctx);
}
