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
package org.zkoss.zss.ui.ua;

import org.zkoss.util.resource.Labels;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.UserActionHandler;
import org.zkoss.zul.Messagebox;

/**
 * @author dennis
 *
 */
public abstract class AbstractUserHandler implements UserActionHandler{

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
}
