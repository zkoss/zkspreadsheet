/* ProtectSheetAction.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/8/5 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ui.sys;

import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.SheetOperationUtil;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.UserActionContext;
import org.zkoss.zss.ui.sys.ua.impl.AbstractSheetHandler;

/**
 * @author dennis
 *
 */
public class ProtectSheetAction extends AbstractSheetHandler {

	/* (non-Javadoc)
	 * @see org.zkoss.zss.ui.sys.ua.impl.AbstractHandler#processAction(org.zkoss.zss.ui.UserActionContext)
	 */
	@Override
	protected boolean processAction(UserActionContext ctx) {
		Sheet sheet = ctx.getSheet();
		
		Range range = Ranges.range(sheet);
		
		String newpassword = getPassword();//TODO, make it meaningful
		if(range.isProtected()){
			SheetOperationUtil.protectSheet(range,null,null);
		}else{
			SheetOperationUtil.protectSheet(range,null,newpassword);
		}
		return true;
	}
	
	protected String getPassword(){
		return "1234";
	}

}
