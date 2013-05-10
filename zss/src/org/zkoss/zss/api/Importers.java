package org.zkoss.zss.api;

import org.zkoss.zss.api.impl.ImporterImpl;
import org.zkoss.zss.model.sys.XImporter;
import org.zkoss.zss.model.sys.XImporters;

public class Importers {

	public static Importer getImporter(String type) {

		XImporter imp = XImporters.getImporter(type);

		return imp == null ? null : new ImporterImpl(imp);
	}
}
