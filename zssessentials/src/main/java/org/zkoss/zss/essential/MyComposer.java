package org.zkoss.zss.essential;

import java.io.IOException;
import java.io.InputStream;

import org.zkoss.util.media.Media;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.event.UploadEvent;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.ui.Spreadsheet;

public class MyComposer extends SelectorComposer<Component> {

	@Wire("spreadsheet")
	Spreadsheet spreadsheet;
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		//initialize stuff here
		//spreadsheet.setSrc("/WEB-INF/books/startzss.xlsx");
	}
	
	
	@Listen("onUpload = button")
	public void showBook(UploadEvent event) throws IOException{
		Media media = event.getMedia();
		if (media.isBinary()) {
			Importer importer = Importers.getImporter();
			InputStream inputStream = WebApps.getCurrent().getResourceAsStream("/WEB-INF/books/startzss.xlsx");
			Book book = importer.imports(inputStream, "startzss");
			spreadsheet.setBook(book);
		}
	}
}
