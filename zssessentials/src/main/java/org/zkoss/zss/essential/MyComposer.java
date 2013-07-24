package org.zkoss.zss.essential;

import java.io.File;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.ui.Spreadsheet;

@SuppressWarnings("serial")
public class MyComposer extends SelectorComposer<Component> {

	@Wire
	Spreadsheet ss;
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);  //wire variables and event listeners
		//access components after calling super.doAfterCompose()
		//ss.setSrc("/WEB-INF/books/startzss.xlsx");
		Importer importer = Importers.getImporter();
		Book book = importer.imports(getFile(), "sample");
		ss.setBook(book);
	}
	
	
	private File getFile() {
		//get a file 
		return new File(WebApps.getCurrent().getRealPath("/WEB-INF/books/startzss.xlsx"));
	}
}
