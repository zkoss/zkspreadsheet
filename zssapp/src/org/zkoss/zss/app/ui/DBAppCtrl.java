package org.zkoss.zss.app.ui;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.app.db.DBDataGrid;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBooks;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NDataGrid;
import org.zkoss.zss.ngmodel.NFont;
import org.zkoss.zss.ngmodel.NFont.Boldweight;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * to demo the capability to load table as a book
 * @author dennis
 *
 */
public class DBAppCtrl extends SelectorComposer<Component> {
	
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
		NCellStyle style = book.createCellStyle(true);
		style.setDataFormat("yyyy/m/d");
		sheet1.getColumn(1).setCellStyle(style);
		sheet1.getColumn(1).setWidth(200);
		style = book.createCellStyle(true);
		style.setDataFormat("###,000");
		NFont font = book.createFont(true);
		font.setBoldweight(Boldweight.BOLD);
		font.setColor(book.createColor("#0000FF"));
		style.setFont(font);
		sheet1.getColumn(3).setCellStyle(style);
		sheet1.getColumn(3).setWidth(200);
		return book;
	}

	private NDataGrid createDataGrid(NSheet sheet) {
		return new DBDataGrid(sheet);
	}
}
