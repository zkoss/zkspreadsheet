/* Step.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 29, 2011 2:24:14 PM, Created by henrichen
}}IS_NOTE

Copyright (C) 2011 Potix Corporation. All Rights Reserved.
*/


package org.zkoss.zss.model.impl;

import org.zkoss.poi.ss.usermodel.Cell;

/**
 * Step for handling incremental auto fill
 * @author henrichen
 * @since 2.1.0
 */
public interface Step {
	public Object next(Cell cell); //return next value of the incremental sequence
}
