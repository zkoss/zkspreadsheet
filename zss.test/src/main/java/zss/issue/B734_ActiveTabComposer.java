/* ExportComposer.java

	Purpose:
		
	Description:
		
	History:
		November 05, 5:53:16 PM     2010, Created by Ashish Dasnurkar

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
 */
package zss.issue;

import java.io.*;
import java.util.Locale;

import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.model.*;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.*;

/**
 * @author hawk
 * 
 */
public class B734_ActiveTabComposer extends GenericForwardComposer {
	Spreadsheet spreadsheet;

	
	public void onClick$exportBtn(Event evt) throws IOException {
		Book wb = spreadsheet.getBook();
		Exporter c = Exporters.getExporter("excel");
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		c.export(wb, baos);
		Filedownload.save(baos.toByteArray(), "application/file",
				wb.getBookName());
	}
	
	public void onClick$setActiveTabBtn(Event evt) {
		Book book = spreadsheet.getBook();
		book.setActiveSheet(book.getSheetIndex(spreadsheet.getSelectedSheet()));
	}
	
}
