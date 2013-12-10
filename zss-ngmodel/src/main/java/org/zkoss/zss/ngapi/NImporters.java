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
package org.zkoss.zss.ngapi;

import java.util.HashMap;
import java.util.LinkedHashMap;

import org.zkoss.zss.ngapi.impl.ExcelImportFactory;
import org.zkoss.zss.ngapi.impl.TestImporterFactory;
/**
 * Get importers by registered name
 * @author dennis
 *
 */
public class NImporters {
	static private HashMap<String,ImporterFactory> factories = new LinkedHashMap<String, ImporterFactory>();
	
	static{
		//default registration
		register("excel",new ExcelImportFactory());
		register("test",new TestImporterFactory());
	}
	
	/**
	 * Get the default importer which is excel format, and it is smart enough to recognize the format(xls or xlsx)
	 * @return
	 */
	static final synchronized public NImporter getImporter(){
		return getImporter("excel");
	}
	
	/**
	 * Get the registered importer
	 * @param name
	 * @return
	 */
	static final synchronized public NImporter getImporter(String name){
		ImporterFactory factory = factories.get(name);
		if(factory!=null){
			return factory.createImporter();
		}
		throw new IllegalStateException("can find any importer named "+name);
	}
	
	static final synchronized public void register(String name,ImporterFactory factory){
		factories.put(name, factory);
	}
}
