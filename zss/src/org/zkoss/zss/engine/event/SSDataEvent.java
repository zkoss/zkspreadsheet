/* SSDataEvent.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 14, 2007 6:08:07 PM, Created by henrichen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/

package org.zkoss.zss.engine.event;

import java.io.Serializable;

import org.zkoss.zk.ui.event.Event;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.model.Book;


/**
 * Defines an event that encapsulates data changes of the spread sheet {@link Book}. 
 *
 * @author henrichen
 */
public class SSDataEvent extends Event implements Serializable {
	private static final long serialVersionUID = 201011250913L;
	/** Identifies one or more changes in the lists contents. */
	public static final String ON_CONTENTS_CHANGE = "onContentsChange";
    /** Identifies the addition of one or more contiguous items to the list. */    
	public static final String ON_RANGE_INSERT = "onRangeInsert";
    /** Identifies the removal of one or more contiguous items from the list. */   
	public static final String ON_RANGE_DELETE = "onRangeDelete";
	/** Identifies the size change of a Cell. */
	public static final String ON_SIZE_CHANGE = "onSizeChange";
	/** Identifyes the associated button change of a Cell. */
	public static final String ON_BTN_CHANGE = "onBtnChange";
	
	/** Identifies the change of a merged range (move or change size). */
	public static final String ON_MERGE_CHANGE = "onMergeChange";
	/** Identifies the add of a merged range. */ 
	public static final String ON_MERGE_ADD = "onMergeAdd";
    /** Identifies the removal of a merged range. */   
	public static final String ON_MERGE_DELETE = "onMergeDelete";
	
	/** Identifies the grid-line status change. */
	public static final String ON_DISPLAY_GRIDLINES = "onDisplayGridlines";
	
	/** Identifies the protect sheet status change. **/
	public static final String ON_PROTECT_SHEET = "onProtectSheet";
	
	/** Indentifies one chart added. **/
	public static final String ON_CHART_ADD = "onChartAdd";
	
	/** Indentifies one picture added. **/
	public static final String ON_PICTURE_ADD = "onPictureAdd";
	
	/** Indentifies the specified picture deleted. **/
	public static final String ON_PICTURE_DELETE = "onPictureRemove";
	
	/** Indentifies the change of a widget. */
	public static final String ON_WIDGET_CHANGE = "onWidgetChange";
	
	/** Identifies no move direction when add or remove a range. */
	public static final int MOVE_NO = 1000;
	/** Identifies move direction as vertical(down or up) when add or remove a range. */
	public static final int MOVE_V = 1001;
	/** Identifies move direction as horizontal(right or left) when add or remove a range. */
	public static final int MOVE_H = 1002;

	private int _direction; //MOVE_NO, MOVE_V, MOVE_H
	private String _password;
	private Ref _rng; //the applied range
	private Ref _org; //the original range
	private Object _payload; //the payload of this event
	
	/**
	 * Constructor of the SSDataEvent.
	 * @param name event name
	 * @param rng a reference to a range of cells.
	 * @param payload payload of this event.
	 */
	public SSDataEvent(String name, Ref rng, Object payload) {
		this(name, rng, null, MOVE_NO);
		_payload = payload;
	}
	
	/**
	 * Constructor of the SSDataEvent.
	 * @param name event name
	 * @param rng a reference to a range of cells.
	 * @param direction direction of the data event(meaningless when type is CONTENTS_CHANGED).
	 */
	public SSDataEvent(String name, Ref rng, int direction) {
		this(name, rng, null, direction);
	}
	
	/**
	 * Constructor of the SSDataEvent.
	 * @param name event name
	 * @param rng a reference to a range of cells.
	 * @param org the original range (used in move/copy/merge, etc.)
	 * @param direction direction of the data event(meaningless when type is CONTENTS_CHANGED).
	 */
	public SSDataEvent(String name, Ref rng, Ref org, int direction) {
		super(name);
		_rng = rng;
		_org = org;
		_direction = direction;
	}
	
	public SSDataEvent(String name, Ref rng, boolean show) {
		super(name);
		_rng = rng;
		_direction = show ? 1 : 0;
	}
	
	public SSDataEvent(String name, Ref rng, String password) {
		super(name);
		_rng = rng;
		_password = password;
	}
	
	public Ref getRef() {
		return _rng;
	}
	
	public Ref getOriginalRef() {
		return _org;
	}
	
	public int getDirection() {
		return _direction;
	}
	
	public boolean isShow() {
		return _direction != 0;
	}
	
	public boolean getProtect() {
		return _password != null;
	}
	
	public String getPassword() {
		return _password;
	}
	
	public Object getPayload() {
		return _payload;
	}
	
	public String toString() {
		return "["+getName()+" -> "+_rng+","+_org+","+_payload+"]";
	}
	
}
