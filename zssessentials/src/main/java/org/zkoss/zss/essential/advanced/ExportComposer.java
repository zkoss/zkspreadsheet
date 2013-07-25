package org.zkoss.zss.essential.advanced;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.zkoss.util.media.AMedia;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.essential.AbstractDemoComposer;
import org.zkoss.zss.essential.util.BookUtil;
import org.zkoss.zul.Filedownload;
import org.zkoss.zul.Radiogroup;

/**
 * This class shows all the public ZK Spreadsheet you can listen to
 * @author dennis
 *
 */
public class ExportComposer extends AbstractDemoComposer {

	private static final long serialVersionUID = 1L;
	
	@Wire
	Radiogroup pdfRadiogroup;
	
	@Wire
	Radiogroup htmlRadiogroup;
	
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
		Filedownload.save(new AMedia(dlname, null, null, file, true));
	}

	@Listen("onClick = #exportPdf")
	public void exportPdf() throws IOException{
		Exporter exporter = Exporters.getExporter("pdf");
		Book book = ss.getBook();
		File file = File.createTempFile(Long.toString(System.currentTimeMillis()),"temp");
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(file);
			
			if(pdfRadiogroup.getSelectedIndex()==1){
//				exporter.export(ss.getSelectedSheet(), fos);
			}else if(pdfRadiogroup.getSelectedIndex()==2){
//				exporter.export(ss.getSelectedSheet(),ss.getSelection(), fos);
			}else{
				exporter.export(book, fos);
			}
			
		}finally{
			if(fos!=null){
				fos.close();
			}
		}
		String dlname = trimDot(book.getBookName())+".pdf";//name for download file
		Filedownload.save(new AMedia(dlname, null, null, file, true));
	}
	
	private String trimDot(String bookName) {
		int index = bookName.lastIndexOf(".");
		return index<0?bookName:bookName.substring(0, index);
	}	
	
	@Listen("onClick = #exportHtml")
	public void exportHtml() throws IOException{
		Exporter exporter = Exporters.getExporter("html");
		Book book = ss.getBook();
		File file = File.createTempFile(Long.toString(System.currentTimeMillis()),"temp");
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(file);
			
			if(htmlRadiogroup.getSelectedIndex()==1){
//				exporter.export(ss.getSelectedSheet(), fos);
			}else if(htmlRadiogroup.getSelectedIndex()==2){
//				exporter.export(ss.getSelectedSheet(),ss.getSelection(), fos);
			}else{
				exporter.export(book, fos);
			}
			
		}finally{
			if(fos!=null){
				fos.close();
			}
		}
		String dlname =  trimDot(book.getBookName())+".html";//name for download file
		Filedownload.save(new AMedia(dlname, null, null, file, true));
	}
}



