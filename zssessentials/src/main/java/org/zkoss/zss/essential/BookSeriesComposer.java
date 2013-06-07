package org.zkoss.zss.essential;

import java.util.Arrays;
import java.util.List;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.BookSeriesBuilder;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Caption;

/**
 * This class shows all the public ZK Spreadsheet you can listen to
 * @author dennis
 *
 */
public class BookSeriesComposer extends AbstractDemoComposer {

	private static final long serialVersionUID = 1L;
	
	@Wire
	Spreadsheet ss2;
	
	@Wire
	Caption book1Cap;
	@Wire
	Caption book2Cap;
	
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		buildBookSeries(ss.getBook());
	}
	
	protected List<String> contirbuteAvailableBooks(){
		return Arrays.asList("book1.xlsx");
	}
	
	protected void buildBookSeries(Book book){
		Book book1 = ss.getBook();
		Book book2 = ss2.getBook();
		
		BookSeriesBuilder.getInstance().buildBookSeries(new Book[]{book1,book2});
		
		book1Cap.setLabel(book1.getBookName());
		book2Cap.setLabel(book2.getBookName());
	}
	
	@Override
	protected void applyBook(Book book){
		ss.setBook(book);
		buildBookSeries(book);
	}
	
	
	@Listen("onClick = #breakSeries")
	public void breakSeries(){
		Book book1 = ss.getBook();
		Book book2 = ss2.getBook();
		BookSeriesBuilder.getInstance().buildBookSeries(new Book[]{book1});
		BookSeriesBuilder.getInstance().buildBookSeries(new Book[]{book2});
		
		ss.invalidate();
		ss2.invalidate();
	}
}



