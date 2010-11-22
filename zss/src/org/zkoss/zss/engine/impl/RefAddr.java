/* RefAddr.java

	Purpose:
		
	Description:
		
	History:
		Nov 19, 2010 4:17:22 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.engine.impl;

import java.io.Serializable;

import org.zkoss.zss.engine.Ref;

/**
 * @author henrichen
 *
 */
public class RefAddr implements Serializable {
	private static final long serialVersionUID = 201005141203L;
	private int _tRow;
	private int _bRow;
	private int _lCol;
	private int _rCol;
	public RefAddr(int tRow, int lCol, int bRow, int rCol) {
		_tRow = tRow;
		_lCol = lCol;
		_bRow = bRow;
		_rCol = rCol;
	}
	public RefAddr(Ref ref) {
		this(ref.getTopRow(), ref.getLeftCol(), ref.getBottomRow(), ref.getRightCol());
	}
	
	//--Object--//
	@Override
	public int hashCode() {
		return (_lCol == _rCol ? _lCol : (_lCol + _rCol)) 
			+ (_tRow == _bRow ? _tRow : (_tRow + _bRow)) << 14; 
	}
	
	@Override
	public boolean equals(Object other) {
		if (this == other) {
			return true;
		}
		if (!(other instanceof RefAddr)) {
			return false;
		}
		final RefAddr o = (RefAddr) other;
		return _lCol == o._lCol && _rCol == o._rCol && _tRow == o._tRow && _bRow == o._bRow;
	}
}

