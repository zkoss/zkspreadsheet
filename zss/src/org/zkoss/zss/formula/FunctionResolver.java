/* FunctionMapperResolver.java

	Purpose:
		
	Description:
		
	History:
		Aug 9, 2010 6:43:35 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.formula;

import org.apache.poi.hssf.record.formula.udf.UDFFinder;
import org.zkoss.xel.FunctionMapper;

/**
 * Interface to glue POI function mechanism to zkoss xel function mechanism.
 * @author henrichen
 *
 */
public interface FunctionResolver {
	/**
	 * Return the associated {@link UDFFinder}.
	 * @return the associated {@link UDFFinder}.
	 */
	public UDFFinder getUDFFinder();
	
	/**
	 * Return the associated {@link FunctionMapper}.
	 * @return the associated {@link FunctionMapper}.
	 */
	public FunctionMapper getFunctionMapper();
}
