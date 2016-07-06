/* TrueMatch.java

	Purpose:
		
	Description:
		
	History:
		May 18, 2016 3:13:53 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.range.impl;

import java.io.Serializable;

import org.zkoss.zss.model.SCell;

/**
 * Always return true or false
 * @author henri
 * @since 3.9.0
 */
public class TrueMatch implements Matchable<SCell>, Serializable {
	private static final long serialVersionUID = -6497660602871191241L;
	
	final private boolean _true;
	public TrueMatch(boolean not) {
		this._true = not;
	}

	@Override
	public boolean match(SCell value) {
		return _true;
	}
}
