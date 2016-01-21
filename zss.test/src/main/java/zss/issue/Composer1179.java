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

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;

import junit.framework.Assert;

import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.BookSeriesBuilder;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range.DeleteShift;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.undo.DeleteCellAction;
import org.zkoss.zss.ui.sys.UndoableAction;
import org.zkoss.zssex.ui.impl.UndoableActionManagerImpl;
import org.zkoss.zul.Window;
import org.zkoss.zk.ui.select.annotation.Listen;


public class Composer1179 extends SelectorComposer<Window> {

    /**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	@Wire
    private Spreadsheet ss;

	@Listen("onClick = #serialize")
	public void serializeSpreadsheet(Event e) throws IOException {
		serialize(ss, "Spreadsheet");
		Sheet sheet = ss.getSelectedSheet();
		UndoableActionManagerImpl um = new UndoableActionManagerImpl();
		UndoableAction action  = new DeleteCellAction("zss.undo.shiftCell",
				sheet, 0, 0, 0, 0, DeleteShift.LEFT);
		um.doAction(action);
		
		serialize(um, "UndoableASctionmanagerImpl");
		alert("Serialize OK");
	}
	
    private void serialize(Object ss, String message) throws IOException {
    	ByteArrayOutputStream os = new ByteArrayOutputStream();
    	ObjectOutputStream oos = new ObjectOutputStream(os);
    	try {
    	    oos.writeObject(ss);
    	} catch (Exception ex) {
    		ex.printStackTrace();
    		Assert.assertFalse("Fail to serialization "+message, true);
    	} finally  {
    		oos.close();
    	}
    	ByteArrayInputStream is = new ByteArrayInputStream(os.toByteArray());
    	ObjectInputStream ois = new ObjectInputStream(is);
    	try {
    		Object ss0 = ois.readObject();
    	} catch (Exception ex) {
    		ex.printStackTrace();
    		Assert.assertFalse("Fail to de-serialization "+message, true);
    	} finally {
    		ois.close();
    	}
    }
}
