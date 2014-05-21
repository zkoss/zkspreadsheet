/* ExportComposer.java

	Purpose:
		
	Description:
		
	History:
		November 05, 5:53:16 PM     2010, Created by Ashish Dasnurkar

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
 */
package zss.issue;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Locale;

import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Window;

/**
 * @author hawk
 * 
 */
public class Composer670 extends GenericForwardComposer<Window> {
	private static final long serialVersionUID = 1L;
	Spreadsheet spreadsheet;
	
	public Composer670(){
		ZssContext.setThreadLocal(new ZssContext(Locale.GERMAN,-1));
	}
	
	public void onClick$exportBtn(Event evt) throws IOException {
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		Book wb = spreadsheet.getBook();
		Exporter c = Exporters.getExporter("excel");
		c.export(wb, baos);
		baos.close();

		ByteArrayInputStream bais = new ByteArrayInputStream(baos.toByteArray());
		Importer i = Importers.getImporter("excel");
		Book wb2 = i.imports(bais, "test");
		spreadsheet.setBook(wb2);
	}
	
	public void onClick$setFormulaWithRange(){
		Ranges.range(spreadsheet.getSelectedSheet(), 2,0).setCellEditText("=A1*0.8");
	}
}
