/* Books.java

	Purpose:
		
	Description:
		
	History:
		Mar 24, 2010 10:47:04 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.impl;

import java.util.HashMap;
import java.util.Map;

import org.zkoss.poi.ss.formula.CollaboratingWorkbooksEnvironment;
import org.zkoss.poi.ss.formula.WorkbookEvaluator;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Books;

/**
 * Represent multiple correlated books.
 * @author henrichen
 *
 */
public class BooksImpl implements Books {
	private Map<String, Book> _books = new HashMap<String, Book>(4);
	
	public BooksImpl(Book[] books) {
		final int len = books.length;
		final String[] workbookNames = new String[len];
		final WorkbookEvaluator[] evaluators = new WorkbookEvaluator[len];
		for(int j = 0; j < books.length; ++j) {
			final Book book = books[j];
			final String bookName = book.getBookName(); 
			workbookNames[j] = bookName;
			evaluators[j] = book.getFormulaEvaluator().getWorkbookEvaluator();
			_books.put(bookName, book);
			BookHelper.setBooks(book, this);
		}
		CollaboratingWorkbooksEnvironment.setup(workbookNames, evaluators);
	}
	
	public Book getBook(String bookName) {
		return _books.get(bookName);
	}
}
