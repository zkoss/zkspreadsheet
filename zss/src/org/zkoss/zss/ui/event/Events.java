/* Events.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Dec 19, 2007 12:48:06 PM     2007, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.event;

import org.zkoss.zss.ui.event.CellEvent;
import org.zkoss.zss.ui.event.CellMouseEvent;
import org.zkoss.zss.ui.event.HeaderMouseEvent;
import org.zkoss.zss.ui.event.StopEditingEvent;


/**
 * @author Dennis.Chen
 *
 */
public class Events {
	
	/** The onCellFocused event (used with {@link CellEvent}).
	 * Sent when cell get focus from client.
	 */
	public static final String ON_CELL_FOUCSED = "onCellFocused";
	
	/** The onStartEditing event (used with {@link CellEvent}).
	 * Sent when cell start editing.
	 */
	public static final String ON_START_EDITING = "onStartEditing";
	
	
	/** The onCellFocused event (used with {@link StopEditingEvent}).
	 * Sent when cell stop editing
	 */
	public static final String ON_STOP_EDITING = "onStopEditing";
	
	/**
	 * The onCellClick event (used with {@link CellMouseEvent}).
	 * Sent when user left click on a cell
	 */
	public static final String ON_CELL_CLICK = "onCellClick";
	
	/**
	 * The onCellRightClick event (used with {@link CellMouseEvent}).
	 * Sent when user right click on a cell
	 */
	public static final String ON_CELL_RIGHT_CLICK = "onCellRightClick";
	
	/**
	 * The onCellDoubleClick event (used with {@link CellMouseEvent}).
	 * Sent when user double click on a cell
	 */
	public static final String ON_CELL_DOUBLE_CLICK = "onCellDoubleClick";
	
	/**
	 * The onFilter event (used with {@link CellMouseEvent}).
	 * Sent whenuser click on the filter button.
	 */
	public static final String ON_FILTER = "onFilter";
	
	/**
	 * The onValidateDrop event (used with {@link CellMouseEvent}).
	 * Sent when user click on the validation drop down button
	 */
	public static final String ON_VALIDATE_DROP = "onValidateDrop";
	
	/** The onCellChange event (used with {@link CellEvent}).
	 * Sent when cell contents changed.
	 */
	public static final String ON_CELL_CHANGE = "onCellChange";
	
	
	/**
	 * The onHeaderClick event (used with {@link HeaderMouseEvent}).
	 * Sent when user left click on a header
	 */
	public static final String ON_HEADER_CLICK = "onHeaderClick";
	
	/**
	 * The onHeaderRightClick event (used with {@link HeaderMouseEvent}).
	 * Sent when user right click on a header
	 */
	public static final String ON_HEADER_RIGHT_CLICK = "onHeaderRightClick";
	
	/**
	 * The onHeaderDoubleClick event (used with {@link HeaderMouseEvent}).
	 * Sent when user double click on a header
	 */
	public static final String ON_HEADER_DOUBLE_CLICK = "onHeaderDoubleClick";
	
	/** 
	 * The onHeaderSzie event (used with {@link HeaderEvent}).
	 * Sent when user resize a header
	 */
	public static final String ON_HEADER_SIZE = "onHeaderSize";
	
	/**
	 * The onCellSelection event (used with {@link CellSelectionEvent}).
	 * Sent when user select a row/column or a range of cells
	 */
	public static final String ON_CELL_SELECTION = "onCellSelection";
	
	/**
	 * The onSelectionChange event (used with {@link SelectionChangeEvent}).
	 * Sent when user move or modify the range of a selection
	 */
	public static final String ON_SELECTION_CHANGE = "onSelectionChange";
	
	/**
	 * The onEditboxEditing event (used with {@link EditboxEditingEvent}).
	 * Sent when user start to typing in the ZSSEditbox
	 */
	public static final String ON_EDITBOX_EDITING = "onEditboxEditing";
	
	/**
	 * The onHyperlink event (used with {@link HyperlinkEvent}).
	 * Sent when user click on the hyperlink of a cell.
	 */
	public static final String ON_HYPERLINK = "onHyperlink";
}
