/* B674Composer.java

	Purpose:
		
	Description:
		
	History:
		Mon, May 19, 2014 10:55:41 AM, Created by RaymondChao

Copyright (C) 2014 Potix Corporation. All Rights Reserved.

*/
package zss.issue;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.model.Exporter;
import org.zkoss.zss.model.Exporters;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.sys.ActionHandler;
import org.zkoss.zul.Messagebox;

/**
 * 
 * @author RaymondChao
 */
public class B674Composer extends SelectorComposer<Component> {
	
	@Wire
	private Spreadsheet ss;
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		ss.setActionHandler(new SaveBookActionHandler());
	}
	
	private class SaveBookActionHandler extends ActionHandler {
		
		public SaveBookActionHandler() {
			super(ss);
		}
		
		@Override
		public void doSaveBook() {
			Spreadsheet spreadsheet = getSpreadsheet();
			if (spreadsheet.getBook() != null) {
				final String filePath = Executions.getCurrent().getDesktop().getWebApp().getRealPath(spreadsheet.getSrc());
		         
				Exporter exporter = Exporters.getExporter("excel");
				FileOutputStream outputStream = null;
				try {
					outputStream = new FileOutputStream(new File(filePath));
					exporter.export(spreadsheet.getBook(), outputStream);
				} catch (FileNotFoundException e) {
					Messagebox.show("Save excel failed");
				} finally {
					if (outputStream != null) {
						try {
							outputStream.close();
						} catch (IOException e) {
						}
					}
				}
			}    
		}
		@Override
		public void doInsertFunction(Rect selection) {
			
		}
		@Override
		public void doColumnWidth(Rect selection) {
			
		}
		@Override
		public void doRowHeight(Rect selection) {
			
		}
		@Override
		public void doNewBook() {
			
		}
		@Override
		public void doExportPDF(Rect selection) {
			
		}
		@Override
		public void doPasteSpecial(Rect selection) {
			
		}
		@Override
		public void doCustomSort(Rect selection) {
			
		}
		@Override
		public void doHyperlink(Rect selection) {
			
		}
		@Override
		public void doFormatCell(Rect selection) {
			
		}
	}
}
