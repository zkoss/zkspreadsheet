package zss.issue;


import java.io.File;

import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zss.api.*;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.*;

public class Composer622 extends SelectorComposer<Window>{

	private static final long serialVersionUID = -257374661283941797L;

	private Spreadsheet ss;
	private Importer importer = Importers.getImporter();

	
	@Override
	public void doAfterCompose(Window comp) throws Exception {
		super.doAfterCompose(comp);
		ss = new Spreadsheet();
		ss.setHflex("1");
		ss.setVflex("1");
		ss.setShowFormulabar(true);
		ss.setShowContextMenu(true);
		ss.setShowToolbar(true);
		ss.setShowSheetbar(true);
		ss.setMaxVisibleRows(100);
		ss.setMaxVisibleColumns(40);
		Book loadedBook = importer.imports(new File(WebApps.getCurrent().getRealPath("/TestFile2007.xlsx")), "TestFile2007.xlsx");
		ss.setBook(loadedBook);
		ss.setParent(comp);
//		ss.afterCompose(); 		
	}

	
	
}
