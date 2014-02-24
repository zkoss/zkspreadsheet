package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NColumnArray;

public abstract class AbstractColumnArrayAdv implements NColumnArray,LinkedModelObject,Serializable{
	private static final long serialVersionUID = 1L;
//	/*package*/ abstract void onModelInternalEvent(ModelInternalEvent event);
	/*package*/ abstract NCellStyle getCellStyle(boolean local);
	/*package*/ abstract void setIndex(int index);
	/*package*/ abstract void setLastIndex(int lastIndex);
}
