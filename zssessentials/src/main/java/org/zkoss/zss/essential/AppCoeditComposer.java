package org.zkoss.zss.essential;

import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.zkoss.zss.api.model.Book;

/**
 * This class shows all the public ZK Spreadsheet you can listen to
 * @author dennis
 *
 */
public class AppCoeditComposer extends CoeditComposer {

	private static final long serialVersionUID = 1L;
	
	static final Map<String,Book> sharedBook = new HashMap<String,Book>();
	
	protected List<String> contirbuteAvailableBooks(){
		return Arrays.asList("full.xlsx");
	}
	
	@Override
	protected Book loadBookFromAvailable(String bookname){
		Book book;
		synchronized (sharedBook){
			book = sharedBook.get(bookname);
			if(book==null){
				book = super.loadBookFromAvailable(bookname);
				if(book!=null){
					book.setShareScope("application");
					sharedBook.put(bookname, book);
				}
			}
		}
		return book;
	}
}



