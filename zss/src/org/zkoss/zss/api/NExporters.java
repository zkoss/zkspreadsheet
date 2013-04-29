package org.zkoss.zss.api;

import org.zkoss.zss.api.impl.NExporterImpl;
import org.zkoss.zss.model.sys.XExporter;
import org.zkoss.zss.model.sys.XExporters;

public class NExporters {

	public static NExporter getExporter(String type) {
		XExporter exp = XExporters.getExporter(type);
		return exp==null?null:new NExporterImpl(exp);
	}

}
