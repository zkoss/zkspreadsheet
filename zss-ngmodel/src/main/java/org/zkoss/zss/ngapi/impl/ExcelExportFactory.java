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
package org.zkoss.zss.ngapi.impl;

import org.zkoss.zss.ngapi.ExporterFactory;
import org.zkoss.zss.ngapi.NExporter;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
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
		if (type == Type.XLSX){
			return new NExcelXlsxExporter();
		}else{
			return new NExcelXlsExporter();
		}
	}
}
