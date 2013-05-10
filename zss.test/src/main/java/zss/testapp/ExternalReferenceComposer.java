package zss.testapp;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.ui.Spreadsheet;

@SuppressWarnings("serial")
public class ExternalReferenceComposer extends SelectorComposer<Component>{
	
	@Wire("#src")
	Spreadsheet srcSpreadsheet;
	@Wire("#dst")
	Spreadsheet dstSpreadsheet;
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
/* FIXME wait new API to rewrite
		//prepare excel importer
		final Importer importer = Importers.getImporter("excel");

		//prepare source book
		final InputStream srcStream = Sessions.getCurrent().getWebApp().getResourceAsStream("/TestRefFile2007.xlsx"); 
		final Book srcbook = importer.imports(srcStream, "TestRefFile2007.xlsx");

		//prepare destination book
		final InputStream dstStream = Sessions.getCurrent().getWebApp().getResourceAsStream("/TestFile2007.xlsx"); 
		final Book dstbook = importer.imports(dstStream, "TestFile2007.xlsx");

		//add both books into a BookSeries
		final Book[] books = new Book[] {srcbook, dstbook}; 
		BookSeries bookSeries = new BookSeriesImpl(books);

		//associate either book to their corresponding UI spreadsheet components
		srcSpreadsheet.setBook(srcbook);
		dstSpreadsheet.setBook(dstbook);
		*/
	}
}
