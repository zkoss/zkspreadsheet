/* FunctionResolverEx.java

	Purpose:
		
	Description:
		
	History:
		Dec 26, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.ngmodel.impl.sys.formula.ex;

import java.lang.reflect.Field;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.zkoss.lang.Library;
import org.zkoss.poi.ss.formula.DependencyTracker;
import org.zkoss.poi.ss.formula.udf.UDFFinder;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.xel.FunctionMapper;
import org.zkoss.xel.XelException;
import org.zkoss.xel.taglib.Taglib;
import org.zkoss.xel.util.TaglibMapper;
import org.zkoss.zss.ngmodel.sys.formula.NFunctionResolver;

/**
 * Enhanced spreadsheet function mapper
 * @author Pao
 */
public class NFunctionResolverEx implements NFunctionResolver {
	public static String EL_FUNCTION_KEY = ""; // should be replace after found UDF finder

	private static final Logger logger = Logger.getLogger(NFunctionResolverEx.class.getName());
	private static final String TAGLIB_KEY = "http://www.zkoss.org/zss/functions";
	private static FunctionMapper _mapper;
	private static UDFFinder _udffinder;
	private static boolean _fail = false;

	static {
		// license check
		// _fail = !org.zkoss.zssex.rt.Runtime.token(Executions.getCurrent()); // TODO zss 3.5
		if(_mapper == null) {
			final String functions = Library.getProperty(TAGLIB_KEY);
			if(functions != null) {
				final TaglibMapper mapper = new TaglibMapper();
				final ClassLocator loc = new ClassLocator();
				String[] libs = functions.split(",");
				for(int j = 0; j < libs.length; ++j) {
					final String lib = libs[j];
					final Taglib taglib = new Taglib("zss", TAGLIB_KEY + "/" + lib.trim());
					try {
						mapper.load(taglib, loc);
					} catch(XelException ex) {
						logger.log(Level.FINE, ex.getMessage(), ex);
						// ignore if cannot find the resource
					}
				}
				_mapper = mapper;
			}
		}
		if(_udffinder == null) {
			try {
				// UDF finder
				Class<?> clazz = Class.forName("org.zkoss.zssex.formula.ZKUDFFinder");
				Field field = clazz.getField("instance");
				_udffinder = (UDFFinder)field.get(null);
				// EL Function Key
				clazz = Class.forName("org.zkoss.zssex.formula.ELEval");
				field = clazz.getField("NAME");
				EL_FUNCTION_KEY = (String)field.get(null);
			} catch(Exception e) {
				logger.log(Level.WARNING, e.getMessage(), e);
			}
		}
	}

	public NFunctionResolverEx() {
	}

	@Override
	public UDFFinder getUDFFinder() {
		return _fail ? null : _udffinder;
	}

	@Override
	public FunctionMapper getFunctionMapper() {
		return _fail ? null : _mapper;
	}

	@Override
	public DependencyTracker getDependencyTracker() {
		return new NDependencyTrackerEx();
	}

}
