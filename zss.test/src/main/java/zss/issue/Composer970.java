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

import java.io.IOException;

import org.zkoss.zk.ui.*;
import org.zkoss.zk.ui.event.EventQueues;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.*;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.ui.Spreadsheet;

public class Composer970 extends SelectorComposer<Component> {

    @Wire
    private Spreadsheet spreadsheet;
    static private Book book;
      
    public void doAfterCompose(Component comp) throws Exception {
        super.doAfterCompose(comp);
        if (book == null) {
        	final Importer importer = Importers.getImporter();
        	try {
        		book = importer.imports(WebApps.getCurrent().getResource("/issue3/book/970.xlsx"),
        				"970.xlsx");
        		book.setShareScope(EventQueues.APPLICATION);
        	} catch (IOException exception) {
        		throw new RuntimeException(exception.getMessage(), exception);
        	}
        }
        spreadsheet.setBook(book);
        spreadsheet.setSelectedSheet(spreadsheet.getBook().getSheetAt(1).getSheetName());
    }
}
