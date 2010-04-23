/* VarRefImpl.java

	Purpose:
		
	Description:
		
	History:
		Apr 8, 2010 10:26:39 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.engine.impl;

import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.engine.RefSheet;

/**
 * Variable Reference
 * @author henrichen
 *
 */
public class VarRefImpl extends AbstractRefImpl implements Ref {
	private final String _name;
	
	/*package*/ VarRefImpl(String name, RefSheet dummy) {
		super(dummy);
		_name = name;
	}
	
	@Override
	public int getBottomRow() {
		return -1;
	}

	@Override
	public int getLeftCol() {
		return -1;
	}
	@Override
	public int getRightCol() {
		return -1;
	}
	@Override
	public int getTopRow() {
		return -1;
	}
	@Override
	public boolean isWholeColumn() {
		return false;
	}
	@Override
	public boolean isWholeRow() {
		return false;
	}
	
	@Override
	public int getColumnCount() {
		return 0;
	}

	@Override
	public int getRowCount() {
		return 0;
	}
	
	//--AbstractRefImpl--//
	@Override
	protected void removeSelf() {
		getOwnerSheet().getOwnerBook().removeVariableRef(_name);
	}

	//--Object--//
	@Override
	public int hashCode() {
		return _name.hashCode();
	}
	@Override
	public boolean equals(Object o) {
		if (o == this) return true;
		if (!(o instanceof VarRefImpl)) return false;
		final VarRefImpl ref = (VarRefImpl) o;
		return _name == null ? _name == ref._name : _name.equals(ref._name);
	}
}
