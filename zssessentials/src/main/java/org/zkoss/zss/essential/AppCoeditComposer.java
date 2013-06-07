package org.zkoss.zss.essential;

import java.io.IOException;

import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
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
	
	protected Book loadBookFromAvailable(String bookname){
		Book book = null;
		
		
		return book;
	}
	
	@Override
	protected void applyBook(Book book){
		book.setShareScope("desktop");
		ss.setBook(book);
		ss2.setBook(book);
	}	
}



