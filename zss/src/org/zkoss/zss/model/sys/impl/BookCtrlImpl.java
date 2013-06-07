/* BookCtrlImpl.java

	Purpose:
		
	Description:
		
	History:
		Dec 7, 2010 11:30:44 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.sys.impl;

import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.WeakHashMap;

import org.zkoss.zss.engine.RefBook;
import org.zkoss.zss.engine.impl.RefBookImpl;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.ui.impl.Focus;

/**
 * Implementation of {@link BookCtrl}.
 * @author henrichen
 *
 */
public class BookCtrlImpl implements BookCtrl {
	private int _shid;
	private int _focusid;
	private WeakHashMap<Object, String> _focusMap = new WeakHashMap<Object, String>(20);
	private Map<String, String> _focusColors = new HashMap<String, String>(20); //id -> Focus
	
	private final static String[] FOCUS_COLORS = 
		new String[]{"#FFC000","#92D050","#00B050","#00B0F0","#0070C0","#002060","#7030A0", "#FFFF00",
					"#4F81BD","#F29436","#9BBB59","#8064A2","#4BACC6","#F79646","#C00000","#FF0000",
					"#0000FF","#008000","#9900CC","#800080","#800000","#FF6600","#CC0099","#00FFFF"};
	
	private int _colorIndex = 0;
	
	@Override
	public synchronized RefBook newRefBook(XBook book) {
		return new RefBookImpl(book.getBookName(), book.getSpreadsheetVersion().getLastRowIndex(), book.getSpreadsheetVersion().getLastColumnIndex());
	}
	
	public synchronized Object nextSheetId() {
		return Integer.toString((_shid++ & 0x7FFFFFFF), 32);
	}

	@Override
	public synchronized String nextFocusId() {
		return Integer.toString((++_focusid & 0x7FFFFFFF), 32);
	}

	private synchronized String nextFocusColor() {
		String color = FOCUS_COLORS[_colorIndex++ % FOCUS_COLORS.length];
		return color;
	}

	@Override
	public synchronized void addFocus(Object focus) {
		if(focus instanceof Focus){
			String id = ((Focus)focus).getId();
			if(!_focusMap.containsKey(focus)){
				String color = _focusColors.get(id);
				if(color==null){
					_focusColors.put(id,color = nextFocusColor());
				}
				((Focus)focus).setColor(color);
			}
			_focusMap.put(focus, ((Focus)focus).getId());
		}else{
			_focusMap.put(focus, focus.toString());
		}
	}

	@Override
	public synchronized void removeFocus(Object focus) {
		_focusMap.remove(focus);
	}

	@Override
	public synchronized boolean containsFocus(Object focus) {
		syncFocus();
		return _focusMap.containsKey(focus);
	}
	
	public synchronized Set<Object> getAllFocus(){
		syncFocus();
		return new HashSet<Object>(_focusMap.keySet()); 
	}
	
	//if browser is closed directly
	private synchronized void syncFocus() { 
		for (final Iterator<Object> it = _focusMap.keySet().iterator(); it.hasNext(); ) {
			Object focus = it.next();
			if(focus instanceof Focus){
				if (((Focus)focus).isDetached()) {
					it.remove();
					_focusColors.remove(((Focus)focus).getId());
				}
			} 
		}
	}
}
