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
import org.zkoss.zk.ui.util.Clients;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.ui.AuxAction;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.UserActionContext;
import org.zkoss.zss.ui.UserActionHandler;
import org.zkoss.zss.ui.UserActionManager;
import org.zkoss.zss.ui.impl.DefaultUserActionManagerCtrl;
import org.zkoss.zul.Messagebox;

/**
 * 
 * @author RaymondChao
 */
public class B674Composer extends SelectorComposer<Component> {
	
	@Wire
	private Spreadsheet spreadsheet;
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		 //initialize custom handlers
        UserActionManager actionManager = spreadsheet.getUserActionManager();
        actionManager.registerHandler(
                DefaultUserActionManagerCtrl.Category.AUXACTION.getName(),
                AuxAction.SAVE_BOOK.getAction(), new SaveBookHandler());
	}
	
	public class SaveBookHandler implements UserActionHandler {
	     
		@Override
		public boolean isEnabled(Book book, Sheet sheet) {
			return book!=null;
		}
		
		@Override
		public boolean process(UserActionContext ctx){
			try{
				Book book = ctx.getBook();
				Exporter exporter = org.zkoss.zss.api.Exporters.getExporter("excel");
				FileOutputStream outputStream = null;
				try {
					final String filePath = Executions.getCurrent().getDesktop().getWebApp().getRealPath(spreadsheet.getSrc());
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
				Clients.showNotification("saved " + book.getBookName(), "info", null, null, 2000, true);
		         
			} catch(Exception e){
				e.printStackTrace();
			}
			return true;
		}
	}
}
