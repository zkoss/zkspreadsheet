/* AreasImpl.java

	Purpose:
		
	Description:
		
	History:
		Oct 28, 2010 3:15:52 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.zkoss.zss.model.Areas;
import org.zkoss.zss.model.Range;

/**
 * Implementation of the {@link Areas}.
 * @author henrichen
 *
 */
public class AreasImpl implements Areas {
	private List<Range> _areas;
	
	/*package*/ AreasImpl() {
		_areas = new ArrayList(4);
	}
	
	@Override
	public int getCount() {
		
		return _areas.size();
	}

	@Override
	public Range getItem() {
		return _areas.get(0);
	}

	@Override
	public Iterator<Range> iterator() {
		return _areas.iterator();
	}
	
	/*package*/ void addArea(Range rng) {
		_areas.add(rng);
	}
}
