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

import java.io.IOException;
import java.io.InputStream;

import org.zkoss.zk.ui.Sessions;
import org.zkoss.zk.ui.util.GenericForwardComposer;
//import org.zkoss.zss.model.sys.XBook;
//import org.zkoss.zss.model.sys.XImporter;
//import org.zkoss.zss.model.sys.XImporters;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.model.Book;
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
		Importer importer = Importers.getImporter("excel");
		final InputStream is = Sessions.getCurrent().getWebApp().getResourceAsStream("/xls/graficas.xlsx");
		Book book;
		try {
			book = importer.imports(is, "graficas.xlsx");
			spreadsheet.setBook(book);
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}

	public void onClick$invalidateSpreadsheetBtn() {
		spreadsheet.invalidate();
	}
}
