package org.zkoss.zss.ngapi.impl;

import org.zkoss.zss.ngapi.ExporterFactory;
import org.zkoss.zss.ngapi.NExporter;

public class ExcelExportFactory implements ExporterFactory{

	public enum Type{
		XLS,XLSX;
	}
	
	Type type;
	
	public ExcelExportFactory(Type type){
		this.type = type;
	}
	
	@Override
	public NExporter createExporter() {
		return new NExcelExporter(type);
	}
}
