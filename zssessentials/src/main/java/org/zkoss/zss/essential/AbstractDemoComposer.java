package org.zkoss.zss.essential;

import java.io.IOException;
import java.util.List;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.AuxAction;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.AuxActionEvent;
import org.zkoss.zul.Grid;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Listbox;

public class AbstractDemoComposer extends SelectorComposer<Component>{
	private static final long serialVersionUID = 1L;

	protected ListModelList<String> infoModel = new ListModelList<String>();
	
	protected ListModelList<String> availableBookModel = new ListModelList<String>();
	
	@Wire
	protected Grid infoList;
	
	@Wire
	private Listbox availableBookList;
	
	@Wire
	protected Spreadsheet ss;
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		initModel();
		if(availableBookList!=null){
			availableBookList.setModel(availableBookModel);
		}
		if(infoList!=null){
			infoList.setModel(infoModel);
		}
		
		addInfo("Spreadsheet initialized");
		
		String bookname = Executions.getCurrent().getParameter("book");
		if(bookname!=null){
			try{
				Book book = loadBookFromAvailable(bookname);
				if(book!=null)
					applyBook(book);
			}catch(Exception x){}
		}
		
		String sheetname = Executions.getCurrent().getParameter("sheet");
		if(sheetname!=null){
			try{
				Sheet ssheet = ss.getBook().getSheet(sheetname);
				if(ssheet!=null){
					ss.setSelectedSheet(ssheet.getSheetName());
				}
			}catch(Exception x){}
		} 
		
		//sync available book selection
		if(ss.getBook()!=null){
			bookname = ss.getBook().getBookName();
			availableBookModel.addToSelection(bookname);
		}
	}

	protected void initModel() {
		availableBookModel.add("blank.xlsx");
		availableBookModel.add("sample.xlsx");
		availableBookModel.add("full.xlsx");
		List<String> c = contirbuteAvailableBooks();
		if(c!=null){
			for(String s:c){
				if(!availableBookModel.contains(s)){
					availableBookModel.add(s);
				}
			}
			
		}
	}
	
	protected List<String> contirbuteAvailableBooks(){
		return null;
	}
	
	protected void addInfo(String info){
		infoModel.add(0, info);
		while(infoModel.size()>100){
			infoModel.remove(infoModel.size()-1);
		}
	}
	
	@Listen("onSelect = #availableBookList")
	public void onBookSelect(){
		String bookname = availableBookModel.getSelection().iterator().next();
		Book book = loadBookFromAvailable(bookname);
		if(book!=null){
			applyBook(book);
		}
	}
	
	protected Book loadBookFromAvailable(String bookname){
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
	protected void applyBook(Book book){
		ss.setBook(book);
	}	
	
	
	@Listen("onClick = #clearInfo")
	public void onClearInfo(){
		infoModel.clear();
	}
	
	@Listen("onAuxAction = #ss")
	public void onAuxActionHandling(AuxActionEvent event){
		//handle extra action when book close
		if(event.getAction().equals(AuxAction.CLOSE_BOOK.toString())){
			availableBookModel.clearSelection();
		}
	}
}
