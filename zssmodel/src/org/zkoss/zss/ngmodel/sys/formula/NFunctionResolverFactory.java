/* FunctionResolverFactory.java

	Purpose:
		
	Description:
		
	History:
		Dec 26, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.ngmodel.sys.formula;

import java.util.logging.Level;
import java.util.logging.Logger;

import org.zkoss.lang.Library;
import org.zkoss.zss.ngmodel.impl.sys.formula.NFunctionResolverImpl;

/**
 * A factory of formula function resolver.
 * @author Pao
 */
public class NFunctionResolverFactory {
	static private Logger logger = Logger.getLogger(NFunctionResolverFactory.class.getName());

	private static Class<?> functionResolverClazz;
	static {
		String clz = Library.getProperty("org.zkoss.zss.model.FunctionResolver.class");
		if(clz!=null){
			try {
				functionResolverClazz = Class.forName(clz);
			} catch(Exception e) {
				logger.log(Level.WARNING, e.getMessage(), e);
			}			
		}
	}

	public static NFunctionResolver createFunctionResolver() {
		try {
			if(functionResolverClazz != null) {
				return (NFunctionResolver)functionResolverClazz.newInstance();
			}
		} catch(Exception e) {
			logger.log(Level.WARNING, e.getMessage(), e);
			functionResolverClazz = null;
		}
		return new NFunctionResolverImpl();
	}
}
