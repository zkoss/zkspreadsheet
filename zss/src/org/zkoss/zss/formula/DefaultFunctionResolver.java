/* SimpleFunctionResolver.java

	Purpose:
		
	Description:
		
	History:
		Aug 10, 2010 12:04:43 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.formula;

import org.zkoss.poi.hssf.record.formula.udf.UDFFinder;
import org.zkoss.poi.ss.formula.DefaultDependencyTracker;
import org.zkoss.poi.ss.formula.DependencyTracker;
import org.zkoss.xel.FunctionMapper;
import org.zkoss.zss.model.Book;

/**
 * Interface to glue POI function to ZK function.
 * @author henrichen
 *
 */
public class DefaultFunctionResolver implements FunctionResolver {

	@Override
	public DependencyTracker getDependencyTracker(Book book) {
		return new DefaultDependencyTracker(book);
	}

	@Override
	public FunctionMapper getFunctionMapper() {
		return null;
	}

	@Override
	public UDFFinder getUDFFinder() {
		return null;
	}
}
