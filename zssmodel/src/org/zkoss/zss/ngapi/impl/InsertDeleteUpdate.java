/* InsertDeleteUpdate.java

	Purpose:
		
	Description:
		
	History:
		Feb 21, 2014 Created by Pao Wang

Copyright (C) 2014 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.ngapi.impl;

import org.zkoss.zss.ngmodel.NSheet;

/**
 * a range of row/column indicates insert/delete changes.
 * @author Pao
 */
public class InsertDeleteUpdate {
	private NSheet sheet;
	private boolean inserted;
	private boolean row;
	private int index;
	private int lastIndex;

	public InsertDeleteUpdate(NSheet sheet, boolean inserted, boolean row, int index, int lastIndex) {
		this.sheet = sheet;
		this.inserted = inserted;
		this.row = row;
		this.index = index;
		this.lastIndex = lastIndex;
	}

	public NSheet getSheet() {
		return sheet;
	}

	public boolean isInserted() {
		return inserted;
	}

	public boolean isRow() {
		return row;
	}

	public int getIndex() {
		return index;
	}

	public int getLastIndex() {
		return lastIndex;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + index;
		result = prime * result + (inserted ? 1231 : 1237);
		result = prime * result + lastIndex;
		result = prime * result + (row ? 1231 : 1237);
		result = prime * result + ((sheet == null) ? 0 : sheet.hashCode());
		return result;
	}

	@Override
	public boolean equals(Object obj) {
		if(this == obj)
			return true;
		if(obj == null)
			return false;
		if(getClass() != obj.getClass())
			return false;
		InsertDeleteUpdate other = (InsertDeleteUpdate)obj;
		if(index != other.index)
			return false;
		if(inserted != other.inserted)
			return false;
		if(lastIndex != other.lastIndex)
			return false;
		if(row != other.row)
			return false;
		if(sheet == null) {
			if(other.sheet != null)
				return false;
		} else if(!sheet.equals(other.sheet))
			return false;
		return true;
	}

}
