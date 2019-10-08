/* AbstractSheetAwareHandler.java

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

import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.*;
import org.zkoss.zss.model.impl.AbstractSheetAdv;
import org.zkoss.zss.ui.UserActionContext;

/**
 * @author dennis
 * @since 3.0.0
 */
public abstract class AbstractBookHandler extends AbstractHandler{
	private static final long serialVersionUID = -8867851919327890720L;

	@Override
	public boolean isEnabled(Book book, Sheet sheet) {
		return book!=null;
	}

	/**
	 * get auto filter of a table at the selection first, then a sheet's auto filter
	 */
	protected SAutoFilter getAutoFilter(UserActionContext context) {
		AbstractSheetAdv internalSheet = (AbstractSheetAdv) context.getSheet().getInternalSheet();
		STable table = internalSheet.getTableByRowCol(context.getSelection().getRow(), context.getSelection().getColumn());

		if (table != null && table.getAutoFilter() != null){
			return table.getAutoFilter();
		}
		return internalSheet.getAutoFilter();
	}
}
