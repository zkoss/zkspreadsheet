/* TableStyleImpl.java

	Purpose:
		
	Description:
		
	History:
		Dec 9, 2014 7:10:04 PM, Created by henrichen

	Copyright (C) 2014 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import org.zkoss.zss.model.STableStyleInfo;

/**
 * Table style
 * @author henri
 * @since 3.8.0
 */
public class TableStyleInfoImpl implements STableStyleInfo {

	private String name;
	private boolean showColumnStripes;
	private boolean showRowStripes;
	private boolean showFirstColumn;
	private boolean showLastColumn;
	
	public TableStyleInfoImpl(String name, boolean showColumnStripes, boolean showRowStrips,
			boolean showFirstColumn, boolean showLastColumn) {
		this.name = name;
		this.showColumnStripes = showColumnStripes;
		this.showRowStripes = showRowStrips;
		this.showFirstColumn = showFirstColumn;
		this.showLastColumn = showLastColumn;
	}
	
	@Override
	public String getName() {
		return name;
	}
	
	@Override
	public void setName(String name) {
		this.name = name;
	}

	@Override
	public boolean isShowColumnStripes() {
		return showColumnStripes;
	}
	
	@Override
	public void setShowColumnStripes(boolean b) {
		showColumnStripes = b;
	}

	@Override
	public boolean isShowRowStripes() {
		return showRowStripes;
	}
	
	@Override
	public void setShowRowStripes(boolean b) {
		showRowStripes = b;
	}

	@Override
	public boolean isShowLastColumn() {
		return showLastColumn;
	}
	
	@Override
	public void setShowLastColumn(boolean b) {
		showLastColumn = b;
	}

	@Override
	public boolean isShowFirstColumn() {
		return showFirstColumn;
	}
	
	@Override
	public void setShowFirstColumn(boolean b) {
		showFirstColumn = b;
	}

}
