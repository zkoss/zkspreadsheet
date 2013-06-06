package org.zkoss.zss.essential;

import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * This class shows all the public ZK Spreadsheet you can listen to
 * @author dennis
 *
 */
public class AppCoeditComposer extends AbstractDemoComposer {

	private static final long serialVersionUID = 1L;
	
	@Wire
	Spreadsheet ss2;
	
	protected void loadBook(Book book){
		book.setShareScope("desktop");
		ss.setBook(book);
		ss2.setBook(book);
	}	
}



