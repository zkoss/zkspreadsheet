package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NColumn;

public abstract class ColumnAdv implements NColumn,LinkedModelObject,Serializable{
	private static final long serialVersionUID = 1L;
	/*package*/ abstract void onModelEvent(ModelEvent event);

}
