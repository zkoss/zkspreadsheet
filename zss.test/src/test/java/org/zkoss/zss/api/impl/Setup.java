package org.zkoss.zss.api.impl;

import org.zkoss.lang.Library;

public class Setup {

	static{
		Library.setProperty("org.zkoss.zss.model.default.Exporter.class", "excel=org.zkoss.zss.model.sys.impl.ExcelExporter");
		Library.setProperty("org.zkoss.zss.model.default.Importer.class", "excel=org.zkoss.zss.model.sys.impl.ExcelImporter");
	}
	
	static void touch(){}
}
