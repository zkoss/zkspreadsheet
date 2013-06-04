package org.zkoss.zss.essential;

import java.io.IOException;

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
import org.zkoss.zss.ui.DefaultUserAction;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.AuxActionEvent;
import org.zkoss.zul.Grid;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Listbox;

public class AbstractDemoComposer extends SelectorComposer<Component>{
	private static final long serialVersionUID = 1L;

	private ListModelList<String> infoModel = new ListModelList<String>();
	
	private ListModelList<String> availableBookModel = new ListModelList<String>();
	
	@Wire
	private Grid infoList;
	
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
		
		String book = Executions.getCurrent().getParameter("book");
		if(book!=null){
			try{
				loadBookFromAvailable(book);
			}catch(Exception x){}
		}
		
		String sheet = Executions.getCurrent().getParameter("sheet");
		if(sheet!=null){
			try{
				Sheet ssheet = ss.getBook().getSheet(sheet);
				if(ssheet!=null){
					ss.setSelectedSheet(ssheet.getSheetName());
				}
			}catch(Exception x){}
		} 
		
		//sync available book selection
		book = ss.getBook().getBookName();
		
		availableBookModel.addToSelection(book);
	}

	private void initModel() {
		availableBookModel.add("blank.xlsx");
		availableBookModel.add("sample.xlsx");
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
		loadBookFromAvailable(bookname);
	}
	
	protected void loadBookFromAvailable(int index){
		String bookname = availableBookModel.get(index);
		loadBookFromAvailable(bookname);
	}
	
	protected void loadBookFromAvailable(String bookname){
		if(!availableBookModel.contains(bookname)){
			return;
		}
		Importer imp = Importers.getImporter();
		try {
			Book book = imp.imports(WebApps.getCurrent().getResource("/WEB-INF/books/"+bookname), bookname);
			ss.setBook(book);
		} catch (IOException e) {
			throw new RuntimeException(e.getMessage(),e);
		}
	}	
	
	
	@Listen("onClick = #clearInfo")
	public void onClearInfo(){
		infoModel.clear();
	}
	
	@Listen("onAuxAction = #ss")
	public void onAuxActionHandling(AuxActionEvent event){
		//handle extra action when book close
		if(event.getAction().equals(DefaultUserAction.CLOSE_BOOK.toString())){
			availableBookModel.clearSelection();
		}
	}
}
