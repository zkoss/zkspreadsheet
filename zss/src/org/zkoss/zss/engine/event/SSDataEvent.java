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

import org.zkoss.zk.ui.event.Event;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Size;


/**
 * Defines an event that encapsulates data changes of the spread sheet {@link Book}. 
 *
 * @author henrichen
 */
public class SSDataEvent extends Event {
	/** Identifies one or more changes in the lists contents. */
	public static final String ON_CONTENTS_CHANGE = "onContentsChange";
    /** Identifies the addition of one or more contiguous items to the list. */    
	public static final String ON_RANGE_INSERT = "onRangeInsert";
    /** Identifies the removal of one or more contiguous items from the list. */   
	public static final String ON_RANGE_DELETE = "onRangeDelete";
	/** Identifies the size change of a Cell. */
	public static final String ON_SIZE_CHANGE = "onSizeChange";
	
	/** Identifies the change of a merged range (move or change size). */
	public static final String ON_MERGE_CHANGE = "onMergeChange";
	/** Identifies the add of a merged range. */ 
	public static final String ON_MERGE_ADD = "onMergeAdd";
    /** Identifies the removal of a merged range. */   
	public static final String ON_MERGE_DELETE = "onMergeDelete";
	
	/** Indentifes the gridline status change. */
	public static final String ON_DISPLAY_GRIDLINES = "onDisplayGridlines";
	
	/** Identifies no move direction when add or remove a range. */
	public static final int MOVE_NO = 1000;
	/** Identifies move direction as vertical(down or up) when add or remove a range. */
	public static final int MOVE_V = 1001;
	/** Identifies move direction as horizontal(right or left) when add or remove a range. */
	public static final int MOVE_H = 1002;

	private int _direction; //MOVE_NO, MOVE_V, MOVE_H
	private Ref _rng; //the applied range
	private Ref _org; //the original range
	private Size _size; //the 
	
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
	
	public SSDataEvent(String name, Ref rng, Size size) {
		super(name);
		_rng = rng;
		_size = size;
		_direction = MOVE_NO;
	}
	
	public SSDataEvent(String name, Ref rng, boolean show) {
		super(name);
		_rng = rng;
		_direction = show ? 1 : 0;
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
	
	public Size getSize() {
		return _size;
	}

	public boolean isShow() {
		return _direction != 0;
	}
	
	public String toString() {
		return "["+getName()+" -> "+_rng+","+_org+"]";
	}
	
}
