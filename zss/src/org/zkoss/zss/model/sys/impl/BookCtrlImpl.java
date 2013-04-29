/* BookCtrlImpl.java

	Purpose:
		
	Description:
		
	History:
		Dec 7, 2010 11:30:44 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.sys.impl;

import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.WeakHashMap;

import org.zkoss.poi.ss.usermodel.PivotCache;
import org.zkoss.poi.ss.util.AreaReference;
import org.zkoss.zss.engine.RefBook;
import org.zkoss.zss.engine.impl.RefBookImpl;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.ui.Focus;

/**
 * Implementation of {@link BookCtrl}.
 * @author henrichen
 *
 */
public class BookCtrlImpl implements BookCtrl {
	private int _shid;
	private int _focusid;
	private WeakHashMap<Object, String> _focusMap = new WeakHashMap<Object, String>(20);
	
	@Override
	public RefBook newRefBook(XBook book) {
		return new RefBookImpl(book.getBookName(), book.getSpreadsheetVersion().getLastRowIndex(), book.getSpreadsheetVersion().getLastColumnIndex());
	}
	
	public Object nextSheetId() {
		return Integer.toString((_shid++ & 0x7FFFFFFF), 32);
	}

	@Override
	public String nextFocusId() {
		return Integer.toString((++_focusid & 0x7FFFFFFF), 32);
	}

	@Override
	public void addFocus(Object focus) {
		_focusMap.put(focus, ((Focus)focus).getId());
	}

	@Override
	public void removeFocus(Object focus) {
		_focusMap.remove(focus);
	}

	@Override
	public boolean containsFocus(Object focus) {
		syncFocus();
		return _focusMap.containsKey(focus);
	}
	
	//if browser is closed directly
	private void syncFocus() { 
		for (final Iterator<Object> it = _focusMap.keySet().iterator(); it.hasNext(); ) {
			final Focus focus = (Focus) it.next();
			if (focus.isDetached()) {
				it.remove();
			}
		}
	}
}
