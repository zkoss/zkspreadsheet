package org.zkoss.zss.essential;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zss.api.model.NBook;
import org.zkoss.zss.api.ui.NSpreadsheet;
import org.zkoss.zss.essential.util.BookUtil;
import org.zkoss.zss.ui.Spreadsheet;

public class MainCtrl extends SelectorComposer<Component> {
	private static final long serialVersionUID = 1L;
	
	NSpreadsheet nss;

	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		Spreadsheet ss = (Spreadsheet)comp.getFellow("ss");
		ss.focus();
		//TODO before this bug fix http://tracker.zkoss.org/browse/ZSS-220, i have to add this go get last selection
		ss.addEventListener("onCellSelection", new EventListener<Event>() {
			public void onEvent(Event event) throws Exception {}
		}); 
		
		ss.setActionHandler(new MainActionHandler());
		
		
		//
		NBook book = readBook();
		nss = new NSpreadsheet(ss);
		nss.setBook(book);
		//
		
	}

	private NBook readBook() {
		// read book for current user and config or request args
		NBook book = null;
		if (true) {// some condition

		}
		if (book == null) {// default from resource
			book = BookUtil.newBook("newBook", NBook.BookType.EXCEL_2007);
//			book = BookUtil.newBook("newBook", NBook.BookType.EXCEL_2003);
		}
		return book;
	}

}
