package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NColumn;

public abstract class AbstractColumn implements NColumn{

	void onModelEvent(ModelEvent event) {
	}

	void release() {
	}

}
