package org.zkoss.zss.ngapi;

import java.util.HashMap;
import java.util.LinkedHashMap;

import org.zkoss.zss.ngapi.impl.ExcelExportFactory;
import org.zkoss.zss.ngapi.impl.ExcelExportFactory.Type;

public class NExporters {

	static private HashMap<String,ExporterFactory> factories = new LinkedHashMap<String, ExporterFactory>();
	
	static{
		//default registration
		register("excel",new ExcelExportFactory(Type.XLSX));
		register("xlsx",new ExcelExportFactory(Type.XLSX));
		register("xls",new ExcelExportFactory(Type.XLS));
	}
	
	static final synchronized public NExporter getExporter(){
		return getExporter("excel");
	}
	
	static final synchronized public NExporter getExporter(String name){
		ExporterFactory factory = factories.get(name);
		return factory==null?null:factory.createExporter();
	}
	
	static final synchronized public void register(String name,ExporterFactory factory){
		factories.put(name, factory);
	}
}
