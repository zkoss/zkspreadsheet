/* IndexableRef.java

	Purpose:
		
	Description:
		
	History:
		Mar 6, 2010 6:25:37 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.engine.impl;

import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.engine.RefSheet;

/**
 * Reference area used with IndexableRefMatrix for hit test.
 * @author henrichen
 *
 */
public class CellRefImpl extends AbstractRefImpl implements Ref, Indexable {
	private final int _tRow;
	private final int _lCol;
	
	public CellRefImpl(int row, int col, RefSheet ownerSheet) {
		super(ownerSheet);
		_tRow = row;
		_lCol = col;
	}
	
	@Override
	public int getLeftCol() {
		return _lCol;
	}

	@Override
	public int getTopRow() {
		return _tRow;
	}

	@Override
	public int getRightCol() {
		return _lCol;
	}

	@Override
	public int getBottomRow() {
		return _tRow;
	}

	@Override
	public boolean isWholeColumn() {
		return false; //A cell impossible to occupy the whole column
	}

	@Override
	public boolean isWholeRow() {
		return false; //A cell impossible to occupy the whole row
	}
	
	@Override
	public int getColumnCount() {
		return 1;
	}

	@Override
	public int getRowCount() {
		return 1;
	}
	
	//--AbstractRefImpl--//
	@Override
	protected void removeSelf() {
		getOwnerSheet().removeRef(_tRow, _lCol, _tRow, _lCol);
	}

	//--Indexable--//
	@Override
	public int getIndex() {
		return getRightCol();
	}

	@Override
	public int compareTo(Indexable o) {
		return o.getIndex() - this.getIndex();
	}
	
	//--Object--//
	public int hashCode() {
		return (_tRow << 14 + _lCol) ^ (getOwnerSheet() == null ? 0 : getOwnerSheet().hashCode());
	}
	
	public boolean equals(Object o) {
		if (o == this) return true;
		if (!(o instanceof CellRefImpl)) return false;
		final CellRefImpl ref = (CellRefImpl) o;
		return ref._tRow == _tRow  
			&& ref._lCol == _lCol
			&& (getOwnerSheet() == null ? 
					getOwnerSheet() == ref.getOwnerSheet() : getOwnerSheet().equals(ref.getOwnerSheet()));
	}
}
