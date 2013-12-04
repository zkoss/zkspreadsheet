package org.zkoss.zss.ngapi;

import java.util.HashMap;
import java.util.LinkedHashMap;

import org.zkoss.zss.ngapi.impl.ExcelImportFactory;

public class NImporters {
static private HashMap<String,ImporterFactory> factories = new LinkedHashMap<String, ImporterFactory>();
	
	static{
		//default registration
		register("excel",new ExcelImportFactory());
	}
	
	static final synchronized public NImporter getImporter(){
		return getImporter("excel");
	}
	
	static final synchronized public NImporter getImporter(String name){
		ImporterFactory factory = factories.get(name);
		return factory==null?null:factory.createImporter();
	}
	
	static final synchronized public void register(String name,ImporterFactory factory){
		factories.put(name, factory);
	}
}
