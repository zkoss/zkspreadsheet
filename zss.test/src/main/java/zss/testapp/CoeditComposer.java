/* CoeditComposer.java

        Purpose:
                
        Description:
                
        History:
                Jan 10, 2012 5:19:59 PM, Created by henri

Copyright (C) 2012 Potix Corporation. All Rights Reserved.
 */
package zss.testapp;

import java.io.InputStream;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Sessions;
import org.zkoss.zk.ui.event.EventQueues;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * @author henri
 */
public class CoeditComposer extends GenericForwardComposer {
	private static Book book = null;
	private static int i = 0;
	private Spreadsheet ss;

	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		synchronized(CoeditComposer.class) {
			if(book == null) { // initialize the shared Book
				final Importer importer = Importers.getImporter("excel");
				final InputStream is = Sessions.getCurrent().getWebApp()
						.getResourceAsStream("/TestFile2007.xlsx");
				book = importer.imports(is, "simple.xlsx");
				book.setShareScope(EventQueues.APPLICATION); // share the work book in Application Scope
			}
		}
		ss.setBook(book);
		ss.setUserName("User "+ ++i);
	}
}
