package org.zkoss.zss.api;

import org.zkoss.zss.model.Exporter;
import org.zkoss.zss.model.Exporters;

public class NExporters {

	public static NExporter getExporter(String type) {
		Exporter exp = Exporters.getExporter(type);
		return exp==null?null:new NExporter(exp);
	}

}
