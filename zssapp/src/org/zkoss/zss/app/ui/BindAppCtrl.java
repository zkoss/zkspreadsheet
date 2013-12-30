package org.zkoss.zss.app.ui;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.app.bind.BindDataGrid;
import org.zkoss.zss.app.db.DBDataGrid;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBooks;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NDataGrid;
import org.zkoss.zss.ngmodel.NFont;
import org.zkoss.zss.ngmodel.NFont.Boldweight;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Filedownload;

/**
 * to demo the capability to load table as a book
 * @author dennis
 *
 */
public class BindAppCtrl extends SelectorComposer<Component> {
	
	@Wire
	Spreadsheet ss;
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		ss.setXBook(createBook());
	}

	private NBook createBook() {
		NBook book = NBooks.createBook("dbbook");
		NSheet sheet1 = book.createSheet("Table Sheet");
		sheet1.setDataGrid(createDataGrid(sheet1));
		
		sheet1.getCell(1, 1).setFormulaValue("Bean1");
		sheet1.getCell(3, 1).setFormulaValue("Bean2");
		
		return book;
	}

	private NDataGrid createDataGrid(NSheet sheet) {
		return new BindDataGrid();
	}
	
	@Listen("onClick = #export")
	public void onExport(){
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		try {
			Exporters.getExporter().export(ss.getBook(), baos);
			ByteArrayInputStream bais = new ByteArrayInputStream(baos.toByteArray());
			Filedownload.save(bais, null,"download.xlsx");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
