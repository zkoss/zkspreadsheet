package org.zkoss.zss.essential;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Listbox;

/**
 * This class shows all the public ZK Spreadsheet you can listen to
 * @author dennis
 *
 */
public class ExportVerifierComposer extends AbstractDemoComposer {

	private static final long serialVersionUID = 1L;
	
	@Wire
	Spreadsheet ss2;
	
	@Listen("onClick = #exportBtn")
	public void doExport() throws IOException{
		Exporter exporter = Exporters.getExporter();
		
		File f = File.createTempFile(Long.toString(System.currentTimeMillis()),"temp");
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(f);
			exporter.export(ss.getBook(), fos);
			fos.close();
			fos = null;//close before read
			
			
			Importer importer = Importers.getImporter("excel");
			//TODO BUG, if book2 has same name as book, then if you move a chart in book , the chart in book2 will also moved.
			Book book2 = importer.imports(f, ss.getBook().getBookName());
			
			ss2.setBook(book2);
			ss2.setSelectedSheet(ss.getSelectedSheet().getSheetName());
		}finally{
			if(fos!=null){
				fos.close();
			}
			f.delete();
		}
	}

}



