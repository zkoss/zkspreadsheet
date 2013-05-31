package test;

import java.util.Date;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.BookSeriesBuilder;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range.DeleteShift;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.ui.Spreadsheet;

public class BookSeriesComposer extends SelectorComposer<Component>{

	@Wire
	Spreadsheet ss1;
	@Wire
	Spreadsheet ss2;
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		Importer importer = Importers.getImporter("excel");
		Book book1 = importer.imports(WebApps.getCurrent().getResource("/WEB-INF/excelsrc/book1.xlsx"), "book1.xlsx");
		Book book2 = importer.imports(WebApps.getCurrent().getResource("/WEB-INF/excelsrc/book2.xlsx"), "book2.xlsx");
		
		BookSeriesBuilder.getInstance().buildBookSeries(new Book[]{book1,book2});
		
		ss1.setBook(book1);
		
		ss2.setBook(book2);
		
	}
	
	@Listen("onClick = button[label='test1']")
	public void test1(){
		Ranges.range(ss2.getSelectedSheet(),"C1").setCellEditText("Changed "+new Date());
	}

	@Listen("onClick = button[label='test2']")
	public void test2(){
		Ranges.range(ss2.getSelectedSheet(),"C1").toColumnRange().delete(DeleteShift.LEFT);
	}
}
