package org.zkoss.zss.ngapi.impl;

import org.zkoss.zss.ngapi.ImporterFactory;
import org.zkoss.zss.ngapi.NImporter;

public class ExcelImportFactory implements ImporterFactory{

	@Override
	public NImporter createImporter() {
		return new NExcelImporter();
	}

}
