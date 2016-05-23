package zss.issue;

import java.io.File;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.event.EventQueues;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.BookSeriesBuilder;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.ui.Spreadsheet;

public class BookReferenceController1223 extends SelectorComposer<Component> {


    @Wire
    private Spreadsheet resumeSpreadsheet;	//profileSpreadsheet
    Importer importer = Importers.getImporter();
    Book resumeBook ;
    Book profileBook;
    
    public void doAfterCompose(Component comp) throws Exception {
        super.doAfterCompose(comp);
    
        profileBook = importer.imports(getFile("/issue3/book/1223-profile.xlsx"),"profile.xlsx");
        resumeBook = importer.imports(getFile("/issue3/book/1223-resume.xlsx"),"resume.xlsx");
//        profileSpreadsheet.setBook(profileBook);
        resumeBook.setShareScope(EventQueues.APPLICATION);
        profileBook.setShareScope(EventQueues.APPLICATION);
        resumeSpreadsheet.setBook(resumeBook);
         
        BookSeriesBuilder.getInstance().buildBookSeries(new Book[]{ resumeBook, profileBook });
    }
 
    private File getFile(String path){
        return new File(WebApps.getCurrent().getRealPath(path));
    }
    
    @Listen("onClick = #switchBook")
    public void switchBook(){
    	if (resumeSpreadsheet.getBook().getInternalBook() == resumeBook.getInternalBook()){
    		resumeSpreadsheet.setBook(profileBook);
    	}else{
    		resumeSpreadsheet.setBook(resumeBook);
    	}
    }
}