package org.zkoss.zss.api;

import org.zkoss.zss.api.impl.ExporterImpl;
import org.zkoss.zss.model.sys.XExporter;
import org.zkoss.zss.model.sys.XExporters;

public class Exporters {

	public static Exporter getExporter(String type) {
		XExporter exp = XExporters.getExporter(type);
		return exp==null?null:new ExporterImpl(exp);
	}

}
