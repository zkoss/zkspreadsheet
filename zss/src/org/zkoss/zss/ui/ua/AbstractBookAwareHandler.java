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
package org.zkoss.zss.ui.ua;

import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

/**
 * @author dennis
 *
 */
public abstract class AbstractBookAwareHandler extends AbstractUserHandler{

	@Override
	public boolean isEnabled(Book book, Sheet sheet) {
		return book!=null;
	}
}
