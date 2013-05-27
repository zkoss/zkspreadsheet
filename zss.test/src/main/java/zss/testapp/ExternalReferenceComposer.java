package zss.testapp;

import java.io.InputStream;
import java.util.LinkedList;
import java.util.List;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Sessions;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.BookSeriesBuilder;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.ui.Spreadsheet;

@SuppressWarnings("serial")
public class ExternalReferenceComposer extends SelectorComposer<Component>{
	
	@Wire("#src")
	Spreadsheet srcSpreadsheet;
	@Wire("#dst")
	Spreadsheet dstSpreadsheet;
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		String year = "2007";
		if (comp.getAttribute("year") != null){
			year = comp.getAttribute("year").toString();
		}
		
		String srcFileName = null;
		String dstFileName = null;
		if ("2007".equals(year)){
			srcFileName = "TestRefFile"+year+".xlsx";
			dstFileName = "TestFile"+year+".xlsx";
		}else{
			srcFileName = "TestRefFile"+year+".xls";
			dstFileName = "TestFile"+year+".xls";
		}
		
		//prepare excel importer
		final Importer importer = Importers.getImporter("excel");

		//prepare source book
		final InputStream srcStream = Sessions.getCurrent().getWebApp().getResourceAsStream("/"+srcFileName); 
		final Book srcBook = importer.imports(srcStream, srcFileName);

		//prepare destination book
		final InputStream dstStream = Sessions.getCurrent().getWebApp().getResourceAsStream("/"+dstFileName); 
		final Book dstBook = importer.imports(dstStream, dstFileName);

		//add both books into a BookSeries
		List<Book> books = new LinkedList<Book>();
		books.add(srcBook);
		books.add(dstBook);
		BookSeriesBuilder.getInstance().buildBookSeries(books.toArray(new Book[0]));

		//associate either book to their corresponding UI spreadsheet components
		srcSpreadsheet.setBook(srcBook);
		dstSpreadsheet.setBook(dstBook);
		
		dstSpreadsheet.setSelectedSheet("cell-reference");
	}
}
