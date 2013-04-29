/* Importers.java

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
 * <p>Provides utility method to get instance of specific {@link XImporter} implementation
 * such as "excel". Other implementations of {@link XImporter} can be registered 
 * via library property in zk.xml
 * </p><p>
 * For example
 * <code>
 * 	<library-property>
 *		<name>org.zkoss.zss.model.Importer.class</name>
 *		<value>csv=org.xyz.com.CSVImporter</value>
 *	</library-property>
 * </code>
 * <p>
 * Use same type name to override default Importer such as ExcelExporter. 
 * Declare multiple type-name=class-name pairs in the library property value 
 * separated by comma.
 * </p> 
 * @author ashish
 *
 */
public class XImporters {

	private static String DEFAULT_ZSS_IMPORTERS_KEY = "org.zkoss.zss.model.default.Importer.class";
	private static String DEFAULT_ZSSEX_IMPORTERS_KEY = "org.zkoss.zssex.model.default.Importer.class";
	private static String USER_DEFINED_IMPORTERS_KEY = "org.zkoss.zss.model.Importer.class";
	private static Map<String,String> typeClss;
	
	/**
	 * Returns instance of specific {@link XImporter} implementation 
	 * as identified by type
	 * @param type
	 * @return Importer instance or null if not found
	 */
	public static XImporter getImporter(String type) {
		if(typeClss == null) {
			typeClss = new HashMap<String,String>();
			loadImporters(DEFAULT_ZSS_IMPORTERS_KEY);
			loadImporters(DEFAULT_ZSSEX_IMPORTERS_KEY);
			loadImporters(USER_DEFINED_IMPORTERS_KEY);
		}
		
		String importerClnm = typeClss.get(type);
		if(importerClnm != null && importerClnm.length() > 0) {
			try {
				Object o = Classes.newInstanceByThread(importerClnm);
				if(o instanceof XImporter) {
					return (XImporter) o;
				}
			} catch(Exception ex) {
			}
		}
		return null;
	}

	private static void loadImporters(String key) {
		String sTypeClss = Library.getProperty(key);
		if(sTypeClss != null) {
			String[] importers = sTypeClss.split(",");
			for(int i=0;i<importers.length;i++) {
				String[] importerClssPair = importers[i].split("=");
				typeClss.put(importerClssPair[0].trim(), importerClssPair[1].trim());
			}
		}
	}
}
