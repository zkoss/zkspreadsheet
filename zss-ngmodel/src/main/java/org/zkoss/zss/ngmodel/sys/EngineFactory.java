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
package org.zkoss.zss.ngmodel.sys;

import org.zkoss.zss.ngmodel.impl.sys.CalendarUtilImpl;
import org.zkoss.zss.ngmodel.impl.sys.DependencyEngineImpl;
import org.zkoss.zss.ngmodel.impl.sys.FormatEngineImpl;
import org.zkoss.zss.ngmodel.impl.sys.FormulaEngineImpl;
import org.zkoss.zss.ngmodel.impl.sys.InputEngineImpl;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyEngine;
import org.zkoss.zss.ngmodel.sys.format.FormatEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.input.InputEngine;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class EngineFactory {

	static private EngineFactory instance;
	
	static private CalendarUtil calendarUtil = new CalendarUtilImpl();

	private EngineFactory() {
	}

	public static EngineFactory getInstance() {
		if (instance == null) {
			synchronized (EngineFactory.class) {
				if (instance == null) {
					// TODO from config
					instance = new EngineFactory();
				}
			}
		}
		return instance;
	}

	public InputEngine createInputEngine() {
		return new InputEngineImpl();
	}

	public FormulaEngine createFormulaEngine() {
		return new FormulaEngineImpl();
	}

	public DependencyEngine createDependencyEngine() {
		return new DependencyEngineImpl();
	}
	
	public FormatEngine createFormatEngine() {
		return new FormatEngineImpl();
	}
	
	public CalendarUtil getCalendarUtil(){
		return calendarUtil;
	}

}
