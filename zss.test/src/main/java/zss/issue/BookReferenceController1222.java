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
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.ui.Spreadsheet;

public class BookReferenceController1222 extends SelectorComposer<Component> {


    @Wire
    private Spreadsheet ss;	
    private Importer importer = Importers.getImporter();
    private int FILE_COUNT = 3;
    private Book[] books = new Book[FILE_COUNT];
    private int index = 0;
    public void doAfterCompose(Component comp) throws Exception {
        super.doAfterCompose(comp);
    
        books[0] = importer.imports(getFile("/issue3/book/1222-dadosbase.xlsx"),"dadosbase.xlsx");
        books[1] = importer.imports(getFile("/issue3/book/1222-pasta1.xlsx"),"pasta1.xlsx");
        books[2] = importer.imports(getFile("/issue3/book/1222-pasta2.xlsx"),"pasta2.xlsx");
//        books[3] = importer.imports(getFile("/WEB-INF/issue/pasta3.xlsx"),"pasta3.xlsx");
//        books[4] = importer.imports(getFile("/WEB-INF/issue/geral.xlsx"),"geral.xlsx");
        
        for (Book b: books){
        	b.setShareScope(EventQueues.APPLICATION);
        }
        ss.setBook(books[index]);
        BookSeriesBuilder.getInstance().buildBookSeries(books);
    }
 
    private File getFile(String path){
        return new File(WebApps.getCurrent().getRealPath(path));
    }
    
    @Listen("onClick = #switchBook")
    public void switchBook(){
    	index = (index+1) % FILE_COUNT;
    	ss.setBook(books[index]);
    }
}