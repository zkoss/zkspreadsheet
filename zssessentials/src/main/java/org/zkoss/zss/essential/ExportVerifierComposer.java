package org.zkoss.zss.essential;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Listbox;

/**
 * This class shows all the public ZK Spreadsheet you can listen to
 * @author dennis
 *
 */
public class ExportVerifierComposer extends SelectorComposer<Component> {

	private static final long serialVersionUID = 1L;
	ListModelList<String> availableBookModel = new ListModelList<String>();
	@Wire
	Listbox availableBookList;
	
	@Wire
	Spreadsheet ss;
	
	@Wire
	Spreadsheet ss2;
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		initModel();
		
		availableBookList.setModel(availableBookModel);;
		
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
	
	@Listen("onClick = #exportBtn")
	public void doExport() throws IOException{
		Exporter exporter = Exporters.getExporter();
		
		File f = File.createTempFile(Long.toString(System.currentTimeMillis()),"temp");
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(f);
			exporter.export(ss.getBook(), fos);
			fos.close();
			fos = null;//close before read
			
			
			Importer importer = Importers.getImporter("excel");
			//TODO BUG, if book2 has same name as book, then if you move a chart in book , the chart in book2 will also moved.
			Book book2 = importer.imports(f, ss.getBook().getBookName());
			
			ss2.setBook(book2);
			ss2.setSelectedSheet(ss.getSelectedSheet().getSheetName());
		}finally{
			if(fos!=null){
				fos.close();
			}
			f.delete();
		}
	}

}



