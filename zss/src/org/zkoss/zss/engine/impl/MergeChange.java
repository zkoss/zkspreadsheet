/* MergeChange.java

	Purpose:
		
	Description:
		
	History:
		May 28, 2010 4:35:32 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.engine.impl;

import org.zkoss.zss.engine.Ref;

/**
 * A pair of reference areas indicate the changes of the merge area. 
 * @author henrichen
 *
 */
public class MergeChange {
	final private Ref _orgMerge; //original merge range
	final private Ref _merge; //merge range changed
	public MergeChange(Ref orgMerge, Ref merge) {
		this._orgMerge = orgMerge;
		this._merge = merge;
	}
	public Ref getOrgMerge() {
		return _orgMerge;
	}

	public Ref getMerge() {
		return _merge;
	}
}
