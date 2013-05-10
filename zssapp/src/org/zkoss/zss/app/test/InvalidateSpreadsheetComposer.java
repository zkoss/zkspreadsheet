/* InvalidateSpreadsheetComposer.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Apr 11, 2012 4:41:57 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.app.test;

import java.io.InputStream;

import org.zkoss.zk.ui.Sessions;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XImporter;
import org.zkoss.zss.model.sys.XImporters;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Button;

/**
 * @author sam
 *
 */
public class InvalidateSpreadsheetComposer extends GenericForwardComposer {

	Spreadsheet spreadsheet;
	Button invalidateSpreadsheetBtn;
	Button setChartBook;
	
	public void onClick$setChartBook() {
		XImporter importer = XImporters.getImporter("excel");
		final InputStream is = Sessions.getCurrent().getWebApp().getResourceAsStream("/xls/graficas.xlsx");
		XBook book = importer.imports(is, "graficas.xlsx");
		spreadsheet.setXBook(book);		
	}

	public void onClick$invalidateSpreadsheetBtn() {
		spreadsheet.invalidate();
	}
}
