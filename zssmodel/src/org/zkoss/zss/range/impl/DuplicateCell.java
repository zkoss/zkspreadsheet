/* Between.java

	Purpose:
		
	Description:
		
	History:
		May 18, 2016 3:13:53 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.range.impl;

import java.io.Serializable;
import java.util.Set;

import org.zkoss.zss.model.SCell;

/**
 * @author henri
 * @since 3.9.0
 */
public class DuplicateCell implements Matchable<SCell>, Serializable {
	private static final long serialVersionUID = -5696013965423062532L;
	
	final private Set<String> cells;
	final private boolean _not;
	
	public DuplicateCell(Set<String> cells, boolean not) {
		this.cells = cells;
		this._not = not;
	}
	@Override
	public boolean match(SCell cell) {
		final String key = ""+cell.getRowIndex()+"_"+cell.getColumnIndex();
		return _not ? !cells.contains(key) : cells.contains(key);
	}
}
