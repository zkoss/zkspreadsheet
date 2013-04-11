package org.zkoss.zss.essential;

import java.io.File;
import java.io.IOException;

import org.zkoss.util.logging.Log;
import org.zkoss.util.media.AMedia;
import org.zkoss.zss.api.model.NBook;
import org.zkoss.zss.api.ui.NSpreadsheet;
import org.zkoss.zss.essential.util.BookUtil;
import org.zkoss.zss.essential.util.ClientUtil;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.sys.ActionHandler;
import org.zkoss.zul.Filedownload;

public class MainActionHandler extends ActionHandler {
	
	private final static Log log = Log.lookup(MainActionHandler.class);
	
	
	
	public NSpreadsheet getNSpreadsheet(){
		return new NSpreadsheet(_spreadsheet);
	}

	@Override
	public void doInsertFunction(Rect selection) {
		ClientUtil.showWarn("Doesn't support yet");

	}

	@Override
	public void doColumnWidth(Rect selection) {
		ClientUtil.showWarn("Doesn't support yet");
	}

	@Override
	public void doRowHeight(Rect selection) {
		ClientUtil.showWarn("Doesn't support yet");
	}

	@Override
	public void doNewBook() {
		NBook book = BookUtil.newBook("newBook", NBook.BookType.EXCEL_2007);
		getNSpreadsheet().setBook(book);
		ClientUtil.showInfo("You are now using a new empty book");
	}

	@Override
	public void doSaveBook() {
		NBook book = getNSpreadsheet().getBook();
		String name = BookUtil.suggestName(book);
		File file;
		try {
			file = BookUtil.saveBook(book);
			Filedownload.save(new AMedia(name, null, "application/vnd.ms-excel", file, true));
		} catch (IOException e) {
			log.error(e.getMessage(),e);
			ClientUtil.showError("Sorry! we can't save the book for you now!");
		}
		
	}

	@Override
	public void doExportPDF(Rect selection) {
		ClientUtil.showWarn("Doesn't support yet");
	}

	@Override
	public void doPasteSpecial(Rect selection) {
		ClientUtil.showWarn("Doesn't support yet");
	}

	@Override
	public void doCustomSort(Rect selection) {
		ClientUtil.showInfo("Doesn't support yet");
	}

	@Override
	public void doHyperlink(Rect selection) {
		ClientUtil.showWarn("Doesn't support yet");
	}

	@Override
	public void doFormatCell(Rect selection) {
		ClientUtil.showWarn("Doesn't support yet");
	}
	
	public void doCloseBook() {
		 getNSpreadsheet().setBook(null);
		super.doCloseBook();
	}

}
