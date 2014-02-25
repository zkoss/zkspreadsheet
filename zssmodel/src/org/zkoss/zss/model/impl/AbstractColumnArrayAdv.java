package org.zkoss.zss.model.impl;

import java.io.Serializable;

import org.zkoss.zss.model.SCellStyle;
import org.zkoss.zss.model.SColumnArray;

public abstract class AbstractColumnArrayAdv implements SColumnArray,LinkedModelObject,Serializable{
	private static final long serialVersionUID = 1L;
//	/*package*/ abstract void onModelInternalEvent(ModelInternalEvent event);
	/*package*/ abstract SCellStyle getCellStyle(boolean local);
	/*package*/ abstract void setIndex(int index);
	/*package*/ abstract void setLastIndex(int lastIndex);
}
