/* InsertDeleteUpdateCollector.java

	Purpose:
		
	Description:
		
	History:
		Feb 21, 2014 Created by Pao Wang

Copyright (C) 2014 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.ngapi.impl;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.zkoss.zss.model.SSheet;

/**
 * A collector for collecting insert/delete row/column updates.
 * @author Pao
 */
public class InsertDeleteUpdateCollector {
	static ThreadLocal<InsertDeleteUpdateCollector> current = new ThreadLocal<InsertDeleteUpdateCollector>();

	private List<InsertDeleteUpdate> insertDeleteUpdates;

	public InsertDeleteUpdateCollector() {
	}

	public static InsertDeleteUpdateCollector setCurrent(InsertDeleteUpdateCollector ctx) {
		InsertDeleteUpdateCollector old = current.get();
		current.set(ctx);
		return old;
	}

	public static InsertDeleteUpdateCollector getCurrent() {
		return current.get();
	}

	public void addInsertDeleteUpdate(SSheet sheet, boolean inserted, boolean isRow, int index, int lastIndex) {
		if(this.insertDeleteUpdates == null) {
			insertDeleteUpdates = new ArrayList<InsertDeleteUpdate>();
		}
		insertDeleteUpdates.add(new InsertDeleteUpdate(sheet, inserted, isRow, index, lastIndex));
	}

	public List<InsertDeleteUpdate> getInsertDeleteUpdates() {
		if(insertDeleteUpdates == null) {
			return Collections.emptyList();
		} else {
			return Collections.unmodifiableList(insertDeleteUpdates);
		}
	}
}
