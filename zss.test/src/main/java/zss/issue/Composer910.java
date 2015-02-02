/**
 * This is about ZSS-910: Problem referencing a cell with formula in same position of other file 
 * @author henrichne
 * @date 2/2/2015
 * 
 * 1. Enter 1 at A1, 2 at B1 in profile.xlsx
 * 2. Enter =SUM(A1:B1) at C1 in profile.xlsx; you should see 3 shown at C1
 * 3. Enter =[profile.xlsx]Source!C1 at C1 in resume.xlsx; you should see 3 shown at C1 of resume.xlsx, too.
 * 4. Delete content at C1 in resume.xlsx
 * 5. Paste the same formula above again, =[profile.xlsx]Source!C1 at C1 in resume.xlsx
 * 6. You should still see 3 shown at C1 in resume.xlsx; it is a bug if not. 
 */

package zss.issue;

import java.io.File;

import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.BookSeriesBuilder;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Window;

public class Composer910 extends SelectorComposer<Window> {

    /**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	@Wire
    private Spreadsheet profileSpreadsheet, resumeSpreadsheet;
    
    public void doAfterCompose(Window comp) throws Exception {
        super.doAfterCompose(comp);
         
        Importer importer = Importers.getImporter();
        Book resumeBook = importer.imports(getFile("/issue3/book/blank.xlsx"),"resume.xlsx");
        Book profileBook = importer.imports(getFile("/issue3/book/blank.xlsx"),"profile.xlsx");
        profileSpreadsheet.setBook(profileBook);
        resumeSpreadsheet.setBook(resumeBook);
         
        BookSeriesBuilder.getInstance().buildBookSeries(new Book[]{ resumeBook, profileBook });
    }
 
    private File getFile(String path){
        return new File(WebApps.getCurrent().getRealPath(path));
    }
}
