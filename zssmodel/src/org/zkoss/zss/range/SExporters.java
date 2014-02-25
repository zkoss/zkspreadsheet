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
package org.zkoss.zss.range;

import java.util.HashMap;
import java.util.LinkedHashMap;

import org.zkoss.zss.range.impl.imexp.ExcelExportFactory;
import org.zkoss.zss.range.impl.imexp.ExcelExportFactory.Type;

/**
 * Get Exporter by registered name
 * @author dennis
 * @sicne 3.5.0
 */
public class SExporters {

	static private HashMap<String,SExporterFactory> factories = new LinkedHashMap<String, SExporterFactory>();
	
	static{
		//default registration
		register("excel",new ExcelExportFactory(Type.XLSX));
		register("xlsx",new ExcelExportFactory(Type.XLSX));
		register("xls",new ExcelExportFactory(Type.XLS));
	}
	
	/**
	 * Gets the default exporter, which is excel xlsx format
	 * @return
	 */
	static final synchronized public SExporter getExporter(){
		return getExporter("excel");
	}
	
	/**
	 * Gets the registered export by name
	 * @param name the exporter name
	 */
	static final synchronized public SExporter getExporter(String name){
		SExporterFactory factory = factories.get(name);
		if(factory!=null){
			return factory.createExporter();
		}
		throw new IllegalStateException("can find any exporter named "+name);
		
	}
	
	/**
	 * Register a exporter factory
	 * @param name the name
	 * @param factory the exporter factory
	 */
	static final synchronized public void register(String name,SExporterFactory factory){
		factories.put(name, factory);
	}
}
