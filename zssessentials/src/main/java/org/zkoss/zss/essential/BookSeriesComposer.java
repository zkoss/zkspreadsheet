package org.zkoss.zss.essential;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.BookSeriesBuilder;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Caption;

/**
 * This class demonstrate usage of BookSeriesBuilder
 * @author dennis, Hawk
 *
 */
@SuppressWarnings("serial")
public class BookSeriesComposer extends SelectorComposer<Component> {

	
	@Wire
	Spreadsheet ss;
	@Wire
	Spreadsheet ss2;
	
	@Wire
	Caption book1Caption;
	@Wire
	Caption book2Caption;
	
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		Book book1 = ss.getBook();
		Book book2 = ss2.getBook();
		
		BookSeriesBuilder.getInstance().buildBookSeries(new Book[]{book1,book2});
		
		book1Caption.setLabel(book1.getBookName());
		book2Caption.setLabel(book2.getBookName());
	}
	
}



