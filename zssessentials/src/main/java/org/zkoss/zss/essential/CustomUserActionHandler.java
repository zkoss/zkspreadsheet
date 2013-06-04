package org.zkoss.zss.essential;

import java.io.File;
import java.io.IOException;
import java.util.Set;

import org.zkoss.util.logging.Log;
import org.zkoss.util.media.AMedia;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.essential.util.BookUtil;
import org.zkoss.zss.essential.util.ClientUtil;
import org.zkoss.zss.ui.DefaultUserAction;
import org.zkoss.zssex.ui.DefaultExUserActionHandler;
import org.zkoss.zul.Filedownload;

public class CustomUserActionHandler extends DefaultExUserActionHandler {
	private static final long serialVersionUID = 1L;
	private final static Log log = Log.lookup(CustomUserActionHandler.class);


	@Override
	public Set<String> getSupportedUserAction(Sheet sheet) {
		Set<String> actions = super.getSupportedUserAction(sheet);
		actions.add(DefaultUserAction.NEW_BOOK.getAction());
		if(sheet==null){
			
			return actions;
		}
		actions.add(DefaultUserAction.SAVE_BOOK.getAction());
		
		return actions;
	}
	
	
	

	@Override
	protected boolean dispatchAction(String action) {
		if(DefaultUserAction.NEW_BOOK.getAction().equals(action)){
			return doNewBook();
		}else if(DefaultUserAction.SAVE_BOOK.getAction().equals(action)){
			return doSaveBook();
		}
		return super.dispatchAction(action);
	}




	public boolean doNewBook() {
//		NBook book = BookUtil.newBook("newBook", NBook.BookType.EXCEL_2003);
		Book book = BookUtil.newBook("newBook", Book.BookType.EXCEL_2007);
		getSpreadsheet().setBook(book);
		ClientUtil.showInfo("You are now using a new empty book");
		return true;
	}

	public boolean doSaveBook() {
		Book book = getSpreadsheet().getBook();
		String name = BookUtil.suggestName(book);
		File file;
		try {
			file = BookUtil.saveBookToTemp(book);
			Filedownload.save(new AMedia(name, null, "application/vnd.ms-excel", file, true));
		} catch (IOException e) {
			log.error(e.getMessage(),e);
			ClientUtil.showError("Sorry! we can't save the book for you now!");
		}
		return true;
	}

//	@Override
//	public boolean doCloseBook() {
//		 getSpreadsheet().setBook(null);
//		return super.doCloseBook();
//	}

}
