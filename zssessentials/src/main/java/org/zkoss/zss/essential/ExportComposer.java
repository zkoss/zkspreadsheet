package org.zkoss.zss.essential;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.zkoss.util.media.AMedia;
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
import org.zkoss.zss.essential.util.BookUtil;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Filedownload;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Listbox;

/**
 * This class shows all the public ZK Spreadsheet you can listen to
 * @author dennis
 *
 */
public class ExportComposer extends SelectorComposer<Component> {

	private static final long serialVersionUID = 1L;
	ListModelList<String> availableBookModel = new ListModelList<String>();
	@Wire
	Listbox availableBookList;
	
	@Wire
	Spreadsheet ss;
	
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
	
	@Listen("onClick = #exportExcel")
	public void doExport() throws IOException{
		Exporter exporter = Exporters.getExporter("excel");
		Book book = ss.getBook();
		File file = File.createTempFile(Long.toString(System.currentTimeMillis()),"temp");
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(file);
			exporter.export(book, fos);
		}finally{
			if(fos!=null){
				fos.close();
			}
		}
		String dlname = BookUtil.suggestName(book);//name for download file
		Filedownload.save(new AMedia(dlname, null, "application/vnd.ms-excel", file, true));
	}

	@Listen("onClick = #exportPdf")
	public void exportPdf() throws IOException{
		Exporter exporter = Exporters.getExporter("pdf");
		Book book = ss.getBook();
		File file = File.createTempFile(Long.toString(System.currentTimeMillis()),"temp");
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(file);
			exporter.export(book, fos);
		}finally{
			if(fos!=null){
				fos.close();
			}
		}
		String dlname = book.getBookName()+".pdf";//name for download file
		Filedownload.save(new AMedia(dlname, null, "application/pdf", file, true));
	}
	
	@Listen("onClick = #exportPdfSheet")
	public void exportPdfSheet() throws IOException{
		Exporter exporter = Exporters.getExporter("pdf");
		Book book = ss.getBook();
		File file = File.createTempFile(Long.toString(System.currentTimeMillis()),"temp");
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(file);
			exporter.export(ss.getSelectedSheet(), fos);
		}finally{
			if(fos!=null){
				fos.close();
			}
		}
		String dlname = book.getBookName()+".pdf";//name for download file
		Filedownload.save(new AMedia(dlname, null, "application/pdf", file, true));
	}
	
	@Listen("onClick = #exportPdfSelection")
	public void exportPdfSelection() throws IOException{
		Exporter exporter = Exporters.getExporter("pdf");
		Book book = ss.getBook();
		File file = File.createTempFile(Long.toString(System.currentTimeMillis()),"temp");
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(file);
			exporter.export(ss.getSelectedSheet(),ss.getSelection(), fos);
		}finally{
			if(fos!=null){
				fos.close();
			}
		}
		String dlname = book.getBookName()+".pdf";//name for download file
		Filedownload.save(new AMedia(dlname, null, "application/pdf", file, true));
	}		
	
	@Listen("onClick = #exportHtml")
	public void exportHtml() throws IOException{
		Exporter exporter = Exporters.getExporter("html");
		Book book = ss.getBook();
		File file = File.createTempFile(Long.toString(System.currentTimeMillis()),"temp");
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(file);
			exporter.export(book, fos);
		}finally{
			if(fos!=null){
				fos.close();
			}
		}
		String dlname = book.getBookName()+".html";//name for download file
		Filedownload.save(new AMedia(dlname, null, "text/html", file, true));
	}
	
	@Listen("onClick = #exportHtmlSheet")
	public void exportHtmlSheet() throws IOException{
		Exporter exporter = Exporters.getExporter("html");
		Book book = ss.getBook();
		File file = File.createTempFile(Long.toString(System.currentTimeMillis()),"temp");
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(file);
			exporter.export(ss.getSelectedSheet(), fos);
		}finally{
			if(fos!=null){
				fos.close();
			}
		}
		String dlname = book.getBookName()+".html";//name for download file
		Filedownload.save(new AMedia(dlname, null, "text/html", file, true));
	}
	
	@Listen("onClick = #exportHtmlSelection")
	public void exportHtmlSelection() throws IOException{
		Exporter exporter = Exporters.getExporter("html");
		Book book = ss.getBook();
		File file = File.createTempFile(Long.toString(System.currentTimeMillis()),"temp");
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(file);
			exporter.export(ss.getSelectedSheet(),ss.getSelection(), fos);
		}finally{
			if(fos!=null){
				fos.close();
			}
		}
		String dlname = book.getBookName()+".html";//name for download file
		Filedownload.save(new AMedia(dlname, null, "text/html", file, true));
	}	
}



