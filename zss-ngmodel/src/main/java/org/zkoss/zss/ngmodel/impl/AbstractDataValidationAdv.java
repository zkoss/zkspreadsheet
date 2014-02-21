package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.NDataValidation;

public abstract class AbstractDataValidationAdv implements NDataValidation,LinkedModelObject,Serializable {
	private static final long serialVersionUID = 1L;
	
	/*package*/ abstract void setRegion(CellRegion region);

	/**
	 * copy data from src , it doesn't copy some essential field, such as region region
	 * @param src
	 */
	/*package*/ abstract void copyFrom(AbstractDataValidationAdv src);
}
