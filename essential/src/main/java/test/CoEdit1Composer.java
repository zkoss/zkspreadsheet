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

public class CoEdit1Composer extends SelectorComposer<Component>{

	@Wire
	Spreadsheet ss1;
	@Wire
	Spreadsheet ss2;
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		Importer importer = Importers.getImporter("excel");
		Book book1 = importer.imports(WebApps.getCurrent().getResource("/WEB-INF/excelsrc/copyPasteBase.xlsx"), "copyPasteBase.xlsx");
		
		//have to set before set to spreadsheet
		book1.setShareScope("desktop");
		
		ss1.setUserName("User1");
		ss2.setUserName("User2");
		
		ss1.setBook(book1);
		ss2.setBook(book1);
		
		
		
	}
	
	@Listen("onClick = button[label='test1']")
	public void test1(){
		Ranges.range(ss2.getSelectedSheet(),"C1").setCellEditText("Changed "+new Date());
	}

	@Listen("onClick = button[label='test2']")
	public void test2(){
		Ranges.range(ss2.getSelectedSheet(),"C1").toColumnRange().delete(DeleteShift.LEFT);
	}
	
	@Listen("onClick = button[label='test3']")
	public void test3(){
		Ranges.range(ss2.getSelectedSheet()).deleteSheet();
	}
}
