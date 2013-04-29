/* AreasImpl.java

	Purpose:
		
	Description:
		
	History:
		Oct 28, 2010 3:15:52 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.sys.impl;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.zkoss.zss.model.sys.XAreas;
import org.zkoss.zss.model.sys.XRange;

/**
 * Implementation of the {@link XAreas}.
 * @author henrichen
 *
 */
public class AreasImpl implements XAreas {
	private List<XRange> _areas;
	
	/*package*/ AreasImpl() {
		_areas = new ArrayList(4);
	}
	
	@Override
	public int getCount() {
		
		return _areas.size();
	}

	@Override
	public XRange getItem() {
		return _areas.get(0);
	}

	@Override
	public Iterator<XRange> iterator() {
		return _areas.iterator();
	}
	
	/*package*/ void addArea(XRange rng) {
		_areas.add(rng);
	}
}
