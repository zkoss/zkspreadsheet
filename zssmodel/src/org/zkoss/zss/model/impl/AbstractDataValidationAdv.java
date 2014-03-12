/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.model.impl;

import java.io.Serializable;

import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SDataValidation;
/**
 * 
 * @author Dennis
 * @since 3.5.0
 */
public abstract class AbstractDataValidationAdv implements SDataValidation,LinkedModelObject,Serializable {
	private static final long serialVersionUID = 1L;
	
	/*package*/ abstract void setRegion(CellRegion region);

	/**
	 * copy data from src , it doesn't copy some essential field, such as region region
	 * @param src
	 */
	/*package*/ abstract void copyFrom(AbstractDataValidationAdv src);
}
