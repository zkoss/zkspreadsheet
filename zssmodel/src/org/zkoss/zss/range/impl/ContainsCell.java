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
public class ContainsCell implements Matchable<SCell>, Serializable {
	
	final private Set<String> cells;
	public ContainsCell(Set<String> cells) {
		this.cells = cells;
	}
	@Override
	public boolean match(SCell cell) {
		final String key = ""+cell.getRowIndex()+"_"+cell.getColumnIndex();
		return cells.contains(key);
	}
}
