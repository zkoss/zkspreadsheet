package org.zkoss.zss.api;

import org.zkoss.zss.api.impl.NImporterImpl;
import org.zkoss.zss.model.sys.XImporter;
import org.zkoss.zss.model.sys.XImporters;

public class NImporters {

	public static NImporter getImporter(String type) {

		XImporter imp = XImporters.getImporter(type);

		return imp == null ? null : new NImporterImpl(imp);
	}
}
