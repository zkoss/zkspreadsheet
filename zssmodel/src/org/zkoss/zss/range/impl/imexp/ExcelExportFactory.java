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
package org.zkoss.zss.range.impl.imexp;

import org.zkoss.zss.range.SExporterFactory;
import org.zkoss.zss.range.SExporter;
/**
 * 
 * @author dennis
 * @author Hawk
 * @since 3.5.0
 */
public class ExcelExportFactory implements SExporterFactory{

	public enum Type{
		XLS,XLSX;
	}
	
	private Type _type;
	
	public ExcelExportFactory(Type type){
		this._type = type;
	}
	
	@Override
	public SExporter createExporter() {
		if (_type == Type.XLSX){
			return new ExcelXlsxExporter();
		}else{
			return new ExcelXlsExporter();
		}
	}
}
