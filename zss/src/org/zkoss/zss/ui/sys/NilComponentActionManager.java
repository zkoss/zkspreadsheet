/* NilUserActionHandler.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/06/03 by Dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
 */
package org.zkoss.zss.ui.sys;

import java.util.Arrays;
import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.Set;

import org.zkoss.zk.ui.event.Event;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
/**
 * The user action handler which provide zero spreadsheet operation handling. <br/>
 * You could set this to a spreadsheet if you want to disable any default operation.
 * 
 * @author dennis
 * @since 3.0.0
 */
public class NilComponentActionManager implements ComponentActionManager {
	
	private static final long serialVersionUID = 1L;
	
	@Override
	public void onEvent(Event event) throws Exception {
		
	}

	@Override
	public Set<String> getInterestedEvents() {
		return null;
	}

	@Override
	public String getCtrlKeys() {
		return null;
	}

	@Override
	public Set<String> getSupportedUserAction(Sheet sheet) {
		return null;
	}

	@Override
	public void bind(Spreadsheet sparedsheet) {
		
	}

	@Override
	public void doAfterLoadBook(Book book) {		
	}
	

}
