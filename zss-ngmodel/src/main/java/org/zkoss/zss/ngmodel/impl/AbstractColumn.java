package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NColumn;

public abstract class AbstractColumn implements NColumn,Serializable{
	private static final long serialVersionUID = 1L;

	/*package*/ void onModelEvent(ModelEvent event) {
	}

	/*package*/ void release() {
	}

}
