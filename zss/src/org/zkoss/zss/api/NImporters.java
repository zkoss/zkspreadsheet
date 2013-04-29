package org.zkoss.zss.api;

import org.zkoss.zss.api.impl.NImporterImpl;
import org.zkoss.zss.model.sys.Importer;
import org.zkoss.zss.model.sys.Importers;

public class NImporters {

	public static NImporter getImporter(String type) {

		Importer imp = Importers.getImporter(type);

		return imp == null ? null : new NImporterImpl(imp);
	}
}
