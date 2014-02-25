package org.zkoss.zss.model.impl;

import java.io.Serializable;

import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SDataValidation;

public abstract class AbstractDataValidationAdv implements SDataValidation,LinkedModelObject,Serializable {
	private static final long serialVersionUID = 1L;
	
	/*package*/ abstract void setRegion(CellRegion region);

	/**
	 * copy data from src , it doesn't copy some essential field, such as region region
	 * @param src
	 */
	/*package*/ abstract void copyFrom(AbstractDataValidationAdv src);
}
