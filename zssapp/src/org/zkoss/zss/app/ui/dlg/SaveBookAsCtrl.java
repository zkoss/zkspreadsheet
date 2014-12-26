/* 
	Purpose:
		
	Description:
		
	History:
		2013/7/10, Created by dennis

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.ui.dlg;

import java.util.Map;

import org.zkoss.lang.Strings;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.app.repository.BookRepository;
import org.zkoss.zss.app.repository.BookRepositoryFactory;
import org.zkoss.zss.app.repository.impl.BookManager;
import org.zkoss.zss.app.repository.impl.BookManagerImpl;
import org.zkoss.zss.app.repository.impl.BookUtil;
import org.zkoss.zss.app.repository.impl.CollaborationInfo;
import org.zkoss.zss.app.repository.impl.SimpleBookInfo;
import org.zkoss.zss.app.ui.UiUtil;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.Textbox;
import org.zkoss.zul.Window;

/**
 * 
 * @author dennis
 *
 */
public class SaveBookAsCtrl extends DlgCtrlBase{
	private static final long serialVersionUID = 1L;
	
	public final static String ARG_NAME = "name";
	public final static String BOOK = "book";
	
	public static final String ON_SAVE = "onSave";
	
	private static BookManager bookManager = BookManagerImpl.getInstance();
	
	@Wire
	Textbox bookName;
	
	BookRepository repo = BookRepositoryFactory.getInstance().getRepository();
		
	private final static String URI = "~./zssapp/dlg/saveBookAs.zul";
	
	public static void show(EventListener<DlgCallbackEvent> callback, String name, Book book) {
		Map arg = newArg(callback);
		arg.put(ARG_NAME, name);
		arg.put(BOOK, book);
		
		Window comp = (Window)Executions.createComponents(URI, null, arg);
		comp.doModal();
		return;
	}

	@Listen("onClick=#save; onOK=#saveAsDlg")
	public void onSave(){
		bookName.clearErrorMessage();
		if(Strings.isBlank(bookName.getValue())){
			bookName.setErrorMessage("empty name is not allowed");
			return;
		}
		
		Book book = (Book) Executions.getCurrent().getArg().get(BOOK);
		final String name = BookUtil.appendExtension(bookName.getValue(), book);
		
		if(bookManager.isBookAttached(new SimpleBookInfo(name))) {			
			String users = CollaborationInfo.getInstance().getUsedUsernames(name);
			bookName.setErrorMessage("Book \"" + name + "\" is in used by " + users + ".");
			return;
		}
		
		String sname = BookUtil.suggestFileName(name, book, repo);
		if(!name.equals(sname)) {
			Messagebox.show("want to overwrite file \"" + name + "\" ?", "ZK Spreadsheet", 
					Messagebox.OK + Messagebox.CANCEL, Messagebox.INFORMATION, new EventListener<Event>() {
				@Override
				public void onEvent(Event event) throws Exception {
					if(event.getData().equals(Messagebox.OK)) {
						postCallback(ON_SAVE, newMap(newEntry(ARG_NAME, name)));
						detach();
					}
				}
			});
		} else {
			postCallback(ON_SAVE, newMap(newEntry(ARG_NAME, name)));
			detach();
		}
	}
	
	@Listen("onClick=#cancel; onCancel=#saveAsDlg")
	public void onCancel(){
		detach();
	}
}
