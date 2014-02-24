/* FunctionResolverFactory.java

	Purpose:
		
	Description:
		
	History:
		Dec 26, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.ngmodel.sys.formula;

import org.zkoss.zss.ngmodel.impl.sys.formula.NFunctionResolverImpl;
import org.zkoss.zss.ngmodel.impl.sys.formula.ex.NFunctionResolverEx;

/**
 * A factory of formula function resolver.
 * @author Pao
 */
public class NFunctionResolverFactory {

	private static boolean EE_EDITION = true; // FIXME zss 3.5

	private final static NFunctionResolver resolver;

	static {
		if(EE_EDITION) {
			resolver = new NFunctionResolverEx();
		} else {
			resolver = new NFunctionResolverImpl();
		}
	}

	public static NFunctionResolver getFunctionResolver() {
		return resolver;
	}
}
