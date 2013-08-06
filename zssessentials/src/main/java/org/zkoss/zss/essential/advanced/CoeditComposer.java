package org.zkoss.zss.essential.advanced;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Textbox;

/**
 * This class shows how to enable collaboration edit.
 * @author dennis, Hawk
 *
 */
public class CoeditComposer extends SelectorComposer<Component> {

	private static final long serialVersionUID = 1L;
	
	@Wire
	private Spreadsheet ss;
	@Wire
	private Listbox availableBookList;
	@Wire
	private Textbox userName;
	
	private ListModelList<String> availableBookModel = new ListModelList<String>();
	
	static private final Map<String,Book> sharedBook = new HashMap<String,Book>();
	static private int userCount = 0;
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		initModel();
		availableBookList.setModel(availableBookModel);
		
		userCount++;
		ss.setUserName("user"+userCount);
	}
	
	private void initModel() {
		availableBookModel.add("blank.xlsx");
		availableBookModel.add("sample.xlsx");
		availableBookModel.add("startzss.xlsx");
	}
	
	@Listen("onSelect = #availableBookList")
	public void onBookSelect(){
		String bookName = availableBookList.getSelectedItem().getValue();
		Book book = loadBookFromAvailable(bookName);
		ss.setBook(book);
	}
	
	private Book loadBookFromAvailable(String bookname){
		Book book;
		synchronized (sharedBook){
			book = sharedBook.get(bookname);
			if(book==null){
				book = importBook(bookname);
				book.setShareScope("application");
				sharedBook.put(bookname, book);
			}
		}
		return book;
	}
	
	@Listen("onClick=#setUserName")
	public void setUserName(){
		ss.setUserName(userName.getValue());
	}
	
	
	private Book importBook(String bookname){
		if(!availableBookModel.contains(bookname)){
			return null;
		}
		Importer imp = Importers.getImporter();
		try {
			Book book = imp.imports(WebApps.getCurrent().getResource("/WEB-INF/books/"+bookname), bookname);
			return book;
		} catch (IOException e) {
			throw new RuntimeException(e.getMessage(),e);
		}
	}
	
}



