package org.zkoss.zss.app.repository.impl;

import java.io.IOException;

import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.app.repository.BookInfo;

public interface BookManager {
	Book readBook(BookInfo info) throws IOException;
	BookInfo updateBook(BookInfo info) throws IOException;
	BookInfo saveBook(BookInfo info, Book book) throws IOException;
	void deleteBook(BookInfo info) throws IOException;
	
	void detachBook(BookInfo info);
	boolean isBookAttached(BookInfo info);
	BookInfo getBookInfo(String bookName);
}