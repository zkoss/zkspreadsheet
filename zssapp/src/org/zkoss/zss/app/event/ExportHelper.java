/* Exporter.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 21, 2010 6:27:01 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.app.event;

import java.util.HashMap;

import org.zkoss.lang.Library;
import org.zkoss.util.resource.Labels;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zss.app.MainWindowCtrl;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Messagebox;

/**
 * A helper class for export file
 * @author Sam
 *
 */
public class ExportHelper {
	private ExportHelper(){}
	
	private final static String KEY_PDF = ExportHelper.class.getSimpleName() + "_ZSSPDF";
	
	public static void onExport(ForwardEvent event) {
		if (!hasZssPdf()) {
			try {
				Messagebox.show("Please download Zss Pdf from ZK");
			} catch (InterruptedException e) {
			}
			return;
		}
		
		String param = (String)event.getData();
		MainWindowCtrl ctrl = MainWindowCtrl.getInstance();
		Spreadsheet ss = ctrl.getSpreadsheet();
		if (param == null || ss == null) {
			return;
		}
		
		if (param.equals(Labels.getLabel("export.pdf"))) {
			//TODO: remove arg, use main
			HashMap arg = new HashMap();
			arg.put("spreadsheet", ss);
			Executions.createComponents("/menus/export/exportToPDF.zul", ctrl.getMainWindow(), arg);
		}
	}
	
	public static boolean hasZssPdf() {
		String val = Library.getProperty(KEY_PDF);
		if (val == null) {
			boolean hasZssPdf = verifyZssPdf();
			Library.setProperty(KEY_PDF, String.valueOf(hasZssPdf));
			return hasZssPdf;
		} else {
			return Boolean.valueOf(Library.getProperty(KEY_PDF));
		}
	}

	/**
	 * Verify whether has zss pdf export function or not
	 */
	private static boolean verifyZssPdf() {
		try {
			Class.forName("org.zkoss.zss.model.impl.PdfExporter");
		} catch (ClassNotFoundException ex) {
			return false;
		}
		return true;
	}
}
