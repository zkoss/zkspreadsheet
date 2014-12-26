package org.zkoss.zss.app.repository.impl;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.zkoss.lang.Strings;
import org.zkoss.zk.ui.event.EventQueues;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.app.repository.BookInfo;
import org.zkoss.zss.app.repository.BookRepository;
import org.zkoss.zss.app.repository.BookRepositoryFactory;

public class BookManagerImpl implements BookManager {

	private BookRepository repo = BookRepositoryFactory.getInstance().getRepository();
	private Map<String, Book> books = new HashMap<String, Book>(repo.list().size() + 5);
	
	@Override
	public Book readBook(BookInfo info) throws IOException {
		synchronized (books) {
			if(books.containsKey(info.getName()))
				return books.get(info.getName());
			
			Book book = repo.load(info);
			book.setShareScope(EventQueues.APPLICATION);
			books.put(info.getName(), book);
			
			return book;
		}
	}

	@Override
	public BookInfo updateBook(BookInfo info) throws IOException {
		synchronized (books) {
			return repo.save(info, books.get(info.getName()));
		}
	}

	@Override
	public BookInfo saveBook(BookInfo info, Book book) throws IOException {
		synchronized (books) {
			if(books.containsKey(info.getName()))
				updateBook(info);
			
			BookInfo newInfo = repo.saveAs(info.getName(), book);
			readBook(newInfo);
			return newInfo;
		}
	}

	@Override
	public void deleteBook(BookInfo info) throws IOException {
		synchronized (books) {
			books.remove(info.getName());
			repo.delete(info);
		}
	}

	@Override
	public void detachBook(BookInfo info) {
		books.remove(info.getName());
	}
	
	@Override
	public boolean isBookAttached(BookInfo info) {
		return books.containsKey(info.getName());
	}
	
	@Override
	public BookInfo getBookInfo(String bookName) {
		BookInfo bookinfo = null;
		if(!Strings.isBlank(bookName)){
			bookName = bookName.trim();
			for(BookInfo info:BookRepositoryFactory.getInstance().getRepository().list()){
				if(bookName.equals(info.getName())){
					bookinfo = info;
					break;
				}
			}
		}
		return bookinfo;
	}
	
	// ======================= Singleton Implementation ======================
	private static BookManagerImpl bookManagerImpl;
	private BookManagerImpl(){};
	public static BookManagerImpl getInstance() {
		if (bookManagerImpl == null)
			bookManagerImpl = new BookManagerImpl();
		return bookManagerImpl;
	}


}
