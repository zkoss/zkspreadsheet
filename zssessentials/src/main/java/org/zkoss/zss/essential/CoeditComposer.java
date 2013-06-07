package org.zkoss.zss.essential;

import java.io.IOException;

import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * This class shows all the public ZK Spreadsheet you can listen to
 * @author dennis
 *
 */
public class CoeditComposer extends AbstractDemoComposer {

	private static final long serialVersionUID = 1L;
	
	@Wire
	Spreadsheet ss2;
	
	protected Book loadBookFromAvailable(String bookname){
		Book book = super.loadBookFromAvailable(bookname);
		if(book!=null){
			book.setShareScope("desktop");
		}
		return book;
	}
	
	@Override
	protected void applyBook(Book book){
		super.applyBook(book);
		ss2.setBook(book);
	}
	
	@Listen("onClick=#detach1")
	public void deatch1(){
		ss.setBook(null);//clean the book registration
		ss.detach();
		ss = null;//clean the reference
	}
	
	@Listen("onClick=#detach2")
	public void detach2(){
		ss2.setBook(null);//clean the book registration
		ss2.detach();
		ss2 = null;//clean the reference
	}
}



