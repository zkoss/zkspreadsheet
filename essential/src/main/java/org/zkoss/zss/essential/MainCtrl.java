package org.zkoss.zss.essential;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.essential.util.BookUtil;
import org.zkoss.zss.ui.Spreadsheet;

public class MainCtrl extends SelectorComposer<Component> {
	private static final long serialVersionUID = 1L;
	
	Spreadsheet nss;

	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		Spreadsheet ss = (Spreadsheet)comp.getFellow("ss");
		ss.focus();
		
		ss.setUserActionHandler(new MainActionHandler());
		
		nss = ss;
		//
		Book book = readBook();
		nss.setBook(book);
		//
		
	}

	private Book readBook() {
		// read book for current user and config or request args
		Book book = null;
		if (true) {// some condition

		}
		if (book == null) {// default from resource
			book = BookUtil.newBook("newBook", Book.BookType.EXCEL_2007);
//			book = BookUtil.newBook("newBook", NBook.BookType.EXCEL_2003);
		}
		return book;
	}

}
