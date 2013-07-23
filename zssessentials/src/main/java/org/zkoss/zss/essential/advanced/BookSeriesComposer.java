package org.zkoss.zss.essential.advanced;

import java.io.File;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.BookSeriesBuilder;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * This class demonstrate usage of BookSeriesBuilder
 * @author dennis, Hawk
 *
 */
@SuppressWarnings("serial")
public class BookSeriesComposer extends SelectorComposer<Component> {
	
	@Wire
	Spreadsheet ss;
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		Importer importer = Importers.getImporter();
		Book book1 = importer.imports(getFile("/WEB-INF/books/resume.xlsx"),"resume.xlsx");
		Book book2 = importer.imports(getFile("/WEB-INF/books/profile.xlsx"),"profile.xlsx");
		//can load more books...
		
		ss.setBook(book1);
		
		BookSeriesBuilder.getInstance().buildBookSeries(new Book[]{book1,book2});
	}
	
	private File getFile(String path){
		return new File(WebApps.getCurrent().getRealPath(path));
	}
}



