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
package org.zkoss.zss.ngapi.impl.imexp;

import org.zkoss.zss.ngapi.ImporterFactory;
import org.zkoss.zss.ngapi.NImporter;
/**
 * 
 * @author dennis
 * @author Hawk
 * @since 3.5.0
 */
public class ExcelImportFactory implements ImporterFactory{

	@Override
	public NImporter createImporter() {
		return new NExcelImportAdapter();
	}

}
