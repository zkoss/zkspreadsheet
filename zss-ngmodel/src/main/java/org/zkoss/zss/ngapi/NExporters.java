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

import org.zkoss.zss.ngapi.impl.ExcelExportFactory;
import org.zkoss.zss.ngapi.impl.ExcelExportFactory.Type;

/**
 * Get Exporter by registered name
 * @author dennis
 * @sicne 3.5.0
 */
public class NExporters {

	static private HashMap<String,ExporterFactory> factories = new LinkedHashMap<String, ExporterFactory>();
	
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
	static final synchronized public NExporter getExporter(){
		return getExporter("excel");
	}
	
	/**
	 * Gets the registered export by name
	 * @param name the exporter name
	 */
	static final synchronized public NExporter getExporter(String name){
		ExporterFactory factory = factories.get(name);
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
	static final synchronized public void register(String name,ExporterFactory factory){
		factories.put(name, factory);
	}
}
