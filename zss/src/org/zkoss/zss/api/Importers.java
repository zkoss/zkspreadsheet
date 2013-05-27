package org.zkoss.zss.api;

import org.zkoss.zss.api.impl.ImporterImpl;
import org.zkoss.zss.model.sys.XImporter;
import org.zkoss.zss.model.sys.XImporters;

/**
 * The main class to get system importer
 * @author dennis
 *
 */
public class Importers {

	/**
	 * Gets importer
	 * @param type the importer type (e.x "excel")
	 * @return exporter instance for the type, null if not found
	 */
	public static Importer getImporter(String type) {

		XImporter imp = XImporters.getImporter(type);

		return imp == null ? null : new ImporterImpl(imp);
	}
}
