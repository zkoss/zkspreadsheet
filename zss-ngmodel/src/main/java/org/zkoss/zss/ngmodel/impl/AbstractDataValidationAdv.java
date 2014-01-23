package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.NDataValidation;

public abstract class AbstractDataValidationAdv implements NDataValidation,LinkedModelObject,Serializable {
	private static final long serialVersionUID = 1L;


	/*package*/ abstract void setRegion(int idx,CellRegion region);
}
