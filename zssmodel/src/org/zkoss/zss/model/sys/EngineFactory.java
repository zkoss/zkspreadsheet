/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.model.sys;

import org.zkoss.lang.Library;
import org.zkoss.util.logging.Log;
import org.zkoss.zss.model.impl.sys.CalendarUtilImpl;
import org.zkoss.zss.model.impl.sys.DependencyTableImpl;
import org.zkoss.zss.model.impl.sys.FormatEngineImpl;
import org.zkoss.zss.model.impl.sys.InputEngineImpl;
import org.zkoss.zss.model.impl.sys.formula.FormulaEngineImpl;
import org.zkoss.zss.model.sys.dependency.DependencyTable;
import org.zkoss.zss.model.sys.format.FormatEngine;
import org.zkoss.zss.model.sys.formula.FormulaEngine;
import org.zkoss.zss.model.sys.input.InputEngine;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class EngineFactory {

	private static final Log _logger = Log.lookup(EngineFactory.class.getName());
	
	static private EngineFactory _instance;
	
	static private CalendarUtil _calendarUtil = new CalendarUtilImpl();

	private EngineFactory() {
	}

	public static EngineFactory getInstance() {
		if (_instance == null) {
			synchronized (EngineFactory.class) {
				if (_instance == null) {
					// TODO from config
					_instance = new EngineFactory();
				}
			}
		}
		return _instance;
	}

	public InputEngine createInputEngine() {
		return new InputEngineImpl();
	}

	static Class<?> formulaEnginClazz;
	static {
		String clz = Library.getProperty("org.zkoss.zss.model.FormulaEngine.class");
		if(clz!=null){
			try {
				formulaEnginClazz = Class.forName(clz);
			} catch(Exception e) {
				_logger.error(e.getMessage(), e);
			}			
		}
		
	}
	
	public FormulaEngine createFormulaEngine() {
		try {
			if(formulaEnginClazz != null) {
				return (FormulaEngine)formulaEnginClazz.newInstance();
			}
		} catch(Exception e) {
			_logger.error(e.getMessage(), e);
			formulaEnginClazz = null;
		}
		return new FormulaEngineImpl();
	}

	public DependencyTable createDependencyTable() {
		return new DependencyTableImpl();
	}
	
	public FormatEngine createFormatEngine() {
		return new FormatEngineImpl();
	}
	
	public CalendarUtil getCalendarUtil(){
		return _calendarUtil;
	}

}
