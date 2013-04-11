package org.zkoss.zss.api.model;

import org.zkoss.zss.model.Importer;
import org.zkoss.zss.model.Importers;

public class NImporters {

	public static NImporter getImporter(String type) {

		Importer imp = Importers.getImporter(type);

		return imp == null ? null : new NImporter(imp);
	}
}
