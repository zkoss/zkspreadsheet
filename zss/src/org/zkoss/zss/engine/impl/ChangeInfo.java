/* ChangeInfo.java

	Purpose:
		
	Description:
		
	History:
		May 28, 2010 4:31:32 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.engine.impl;

import java.util.List;
import java.util.Set;

import org.zkoss.zss.engine.Ref;

/**
 * The change information caused by an operation to the spreadsheet sheet/book.
 * @author henrichen
 *
 */
public class ChangeInfo {
	final private Set<Ref> _affected; //all affected reference area
	final private Set<Ref> _toEval; //to be evaluated reference area that will recursively evaluate all affected reference area
	final private List<MergeChange> _mergeChanges; //list of changed merge ranges
	
	public ChangeInfo(Set<Ref> toEval, Set<Ref> affected, List<MergeChange> mergeChanges) {
		this._affected = affected;
		this._toEval = toEval;
		this._mergeChanges = mergeChanges;
	}
	
	public Set<Ref> getAffected() {
		return _affected;
	}

	public Set<Ref> getToEval() {
		return _toEval;
	}

	public List<MergeChange> getMergeChanges() {
		return _mergeChanges;
	}
}
