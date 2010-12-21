/* Exporter.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 21, 2010 6:27:01 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.event;

import java.util.HashMap;

import org.zkoss.lang.Library;
import org.zkoss.util.resource.Labels;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zss.app.MainWindowCtrl;
import org.zkoss.zss.app.zul.Zssapps;
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

	public static void doExportToPDF(Spreadsheet spreadsheet) {
		if (!hasZssPdf()) {
			try {
				Messagebox.show("Please download Zss Pdf from ZK");
			} catch (InterruptedException e) {
			}
			return;
		}
		Executions.createComponents("~./zssapp/html/dialog/exportToPDF.zul", null, Zssapps.newSpreadsheetArg(spreadsheet));
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
			Class.forName("org.zkoss.zss.model.impl.pdf.PdfExporter");
		} catch (ClassNotFoundException ex) {
			return false;
		}
		return true;
	}
}
