/* Exporters.java

	Purpose:
		
	Description:
		
	History:
		Dec 13, 2010 11:06:01 PM, Created by ashish

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.sys;

import java.util.HashMap;
import java.util.Map;

import org.zkoss.lang.Classes;
import org.zkoss.lang.Library;

/**
 * <p>Provides utility method to get instance of specific {@link XExporter} implementation
 * such as "excel" or "pdf". Other implementations of {@link XExporter} can be registered 
 * via library property in zk.xml
 * </p><p>
 * For example
 * <code>
 * 	<library-property>
 *		<name>org.zkoss.zss.model.Exporter.class</name>
 *		<value>csv=org.xyz.com.CSVExporter</value>
 *	</library-property>
 * </code>
 * <p>
 * Use same type-name to override default Importer such as ExcelExporter. 
 * Declare multiple type-name=class-name pairs in the library property 
 * value separated by comma.
 * </p> 
 * @author ashish
 *
 */
public class XExporters {
	private static String DEFAULT_ZSS_EXPORTERS_KEY = "org.zkoss.zss.model.default.Exporter.class";
	private static String DEFAULT_ZSSEX_EXPORTERS_KEY = "org.zkoss.zssex.model.default.Exporter.class";
	private static String USER_DEFINED_EXPORTERS_KEY = "org.zkoss.zss.model.Exporter.class";
	private static Map<String,String> typeClss;
	
	/**
	 * Returns specific (@link Exporter} implementation as identified by type.
	 * @param type
	 * @return Exporter instance or null if not found
	 */
	public static XExporter getExporter(String type) {
		if(typeClss == null) {
			typeClss = new HashMap<String,String>();
			loadExporters(DEFAULT_ZSS_EXPORTERS_KEY);
			loadExporters(DEFAULT_ZSSEX_EXPORTERS_KEY);
			loadExporters(USER_DEFINED_EXPORTERS_KEY);
		}
		
		String exporterClnm = typeClss.get(type);
		if(exporterClnm != null && exporterClnm.length() > 0) {
			try {
				Object o = Classes.newInstanceByThread(exporterClnm);
				if(o instanceof XExporter) {
					return (XExporter) o;
				}
			} catch(Exception ex) {
			}
		}
		return null;
	}

	private static void loadExporters(String key) {
		String sTypeClss = Library.getProperty(key);
		if(sTypeClss != null) {
			String[] exporters = sTypeClss.split(",");
			
			for(int i=0;i<exporters.length;i++) {
				String[] exporterClssPair = exporters[i].split("=");
				typeClss.put(exporterClssPair[0].trim(), exporterClssPair[1].trim());
			}
		}
	}
}
