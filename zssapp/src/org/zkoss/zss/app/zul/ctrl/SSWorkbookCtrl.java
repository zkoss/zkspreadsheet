/* SSWorkbookCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 25, 2010 10:53:10 AM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul.ctrl;

import java.io.ByteArrayOutputStream;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.List;

import org.zkoss.util.media.Media;
import org.zkoss.zk.ui.Desktop;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.util.Clients;
import org.zkoss.zss.app.cell.CellHelper;
import org.zkoss.zss.app.cell.EditHelper;
import org.zkoss.zss.app.file.FileHelper;
import org.zkoss.zss.app.file.SpreadSheetMetaInfo;
import org.zkoss.zss.app.sheet.SheetHelper;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Exporter;
import org.zkoss.zss.model.Exporters;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.Widget;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zss.ui.sys.SpreadsheetCtrl;
import org.zkoss.zul.Image;
import org.zkoss.zul.Messagebox;

/**
 * 
 * @author Sam
 *
 */
public class SSWorkbookCtrl implements WorkbookCtrl {

	private Spreadsheet spreadsheet;
	
	private String lastSheetName = null;
	private HashMap<String, List<Widget>> sheetWidgets = new HashMap<String, List<Widget>>(); 

	/* book event listeners; Boolean value means the listener has subscribed on book or not */
	private HashMap<EventListener, Boolean> bookListeners = new HashMap<EventListener, Boolean>();
	
	/* key to access all books in desktop */
	private final static String KEY_DESKTOP_BOOKS = "org.zkoss.zss.app.zul.ctrl.desktopBooks";
	
	public SSWorkbookCtrl(Spreadsheet spreadsheet) {
		this.spreadsheet = spreadsheet;
	}
	
	public void clearSelectionContent() {
		if (spreadsheet.getSelection() == null)
			return;
		
		CellHelper.clearContent(spreadsheet, SheetHelper.getSpreadsheetMaxSelection(spreadsheet));
	}

	public void clearSelectionStyle() {
		if (spreadsheet.getSelection() == null)
			return;
		
		CellHelper.clearStyle(spreadsheet, SheetHelper.getSpreadsheetMaxSelection(spreadsheet));
	}

	public void copySelection() {
		EditHelper.doCopy(spreadsheet);
	}


	public void cutSelection() {
		EditHelper.doCut(spreadsheet);
	}
	
	public void pasteSelection() {
		EditHelper.doPaste(spreadsheet);
	}

	public void insertColumnLeft() {
		if (spreadsheet.getSelection() == null)
			return;
		
		Rect rect = spreadsheet.getSelection();
		CellHelper.shiftEntireColumnRight(spreadsheet.getSelectedSheet(), 
				rect.getLeft(), rect.getRight());
	}

	public void deleteColumn() {
		Rect rect = spreadsheet.getSelection();
		if (rect == null)
			return;
		
		CellHelper.shiftEntireColumnLeft(spreadsheet.getSelectedSheet(), 
				rect.getLeft(), rect.getRight());
	}

	public void insertRowAbove() {
		CellHelper.shiftEntireRowDown(spreadsheet.getSelectedSheet(), 
				spreadsheet.getSelection().getTop(), 
				spreadsheet.getSelection().getBottom());
	}

	public void deleteRow() {
		Rect rect = spreadsheet.getSelection();
		CellHelper.shiftEntireRowUp(spreadsheet.getSelectedSheet(), rect.getTop(), rect.getBottom());
	}


	public void setSelectedSheet(String name) {
		//TODO: remove last sheet widget shall not handle by AP
		Worksheet lastsheet = spreadsheet.getSelectedSheet();
		List<Widget> rmWgtList = sheetWidgets.get(lastSheetName);
		if (rmWgtList != null) {
			SpreadsheetCtrl ctrl = (SpreadsheetCtrl) spreadsheet.getExtraCtrl();
			for (Widget w : rmWgtList) {
				ctrl.removeWidget(w);
			}
		}
		
		Integer lastIdx = Integer.valueOf(spreadsheet.getBook().getSheetIndex(lastsheet));
		if (lastIdx < 0) //sheet deleted
			sheetWidgets.remove(lastSheetName);
		
		spreadsheet.setSelectedSheet(name);		
		//handle the copy/cut highlight
		final Worksheet sheet = EditHelper.getSourceSheet(spreadsheet);
		if (sheet != null) {
			if (sheet.equals(spreadsheet.getSelectedSheet())) {
				spreadsheet.setHighlight(EditHelper.getSourceRange(spreadsheet));
			} else {
				spreadsheet.setHighlight(null);
			}
		}
		
		//TODO: insert sheet widget shall not handle by AP
		lastSheetName = name;
		List<Widget> addWgtList = sheetWidgets.get(lastSheetName);
		if (addWgtList != null) {
			SpreadsheetCtrl ctrl = (SpreadsheetCtrl) spreadsheet.getExtraCtrl();
			for (Widget w : addWgtList) {
				ctrl.addWidget(w);
			}
		}
		
		spreadsheet.focus(); //move focus in to spreadsheet
	}

	public void hide(boolean hide) {
		Rect rect = spreadsheet.getSelection();
		if (rect == null)
			return;
		
		Ranges.range(spreadsheet.getSelectedSheet(), 
				rect.getTop(), rect.getLeft(), 
				rect.getBottom(), rect.getRight()).setHidden(hide);
	}

	public void insertImage(Media media) {
		//TODO: insert image shall not handle by AP
		try {
			if (media instanceof org.zkoss.image.Image && WebApps.getFeature("pe")) {
				Class imageWidget = null;
				try {
					imageWidget = Class.forName("org.zkoss.zssex.ui.widget.ImageWidget");
				} catch (ClassNotFoundException ex) {
					return;
				}
				Widget widget = (Widget)imageWidget.newInstance();
				Method setImageMethod = imageWidget.getDeclaredMethod("setContent", org.zkoss.image.Image.class);
				setImageMethod.invoke(widget, (Image)media);
				widget.setRow(spreadsheet.getSelection().getTop());
				widget.setColumn(spreadsheet.getSelection().getLeft());
				
				Worksheet seldSheet = spreadsheet.getSelectedSheet();
				String sheetName = seldSheet.getSheetName();
				List<Widget> wgtList = sheetWidgets.get(sheetName);
				if (wgtList == null)
					sheetWidgets.put(sheetName, wgtList = new ArrayList<Widget>());
				wgtList.add(widget);
				
				SpreadsheetCtrl ctrl = (SpreadsheetCtrl) spreadsheet.getExtraCtrl();
				ctrl.addWidget(widget);
			} else if (media != null) {
				Messagebox.show("Not an image: " + media, "Error",
						Messagebox.OK, Messagebox.ERROR);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void insertSheet() {
		int sheetCount = spreadsheet.getBook().getNumberOfSheets();
		spreadsheet.getBook().createSheet("sheet " + (sheetCount + 1));
	}

	public void reGainFocus() {
		Clients.evalJavaScript("zk.Widget.$('" + spreadsheet.getUuid() + "').focus(false);");
	}

	public void renameSelectedSheet(String name) {
		SheetHelper.renameSheet(spreadsheet, name);
		List<Widget> wgts = sheetWidgets.get(lastSheetName);
		if (wgts != null) {
			sheetWidgets.remove(lastSheetName);
			sheetWidgets.put(name, wgts);
			SpreadsheetCtrl ctrl = (SpreadsheetCtrl) spreadsheet.getExtraCtrl();
			for (Widget w : wgts)
				ctrl.addWidget(w);
		}
	}

	public void shiftCell(int direction) {
		Worksheet sheet = spreadsheet.getSelectedSheet();
		Rect rect = spreadsheet.getSelection();
		
		switch (direction) {
		case DesktopWorkbenchContext.SHIFT_CELL_UP:
			CellHelper.shiftCellUp(sheet, rect); 
			break;
		case DesktopWorkbenchContext.SHIFT_CELL_RIGHT:
			CellHelper.shiftCellRight(sheet, rect);
			break;
		case DesktopWorkbenchContext.SHIFT_CELL_DOWN:
			CellHelper.shiftCellDown(sheet, rect); 
			break;
		case DesktopWorkbenchContext.SHIFT_CELL_LEFT:
			CellHelper.shiftCellLeft(sheet, rect); 
			break;
		}
	}

	public void sort(boolean isSortDescending) {
		if (isSortDescending)
			CellHelper.sortDescending(spreadsheet.getSelectedSheet(), SheetHelper.getSpreadsheetMaxSelection(spreadsheet));
		else
			CellHelper.sortAscending(spreadsheet.getSelectedSheet(), SheetHelper.getSpreadsheetMaxSelection(spreadsheet));
	}

	public void setColumnFreeze(int columnFreeze) {
		spreadsheet.setColumnfreeze(columnFreeze);
	}

	public void setRowFreeze(int rowfreeze) {
		spreadsheet.setRowfreeze(rowfreeze);
	}

	public void insertFormula(String formula) {
		Rect rect = spreadsheet.getSelection();
		Range rng = Ranges.range(spreadsheet.getSelectedSheet(), rect.getTop(), rect.getLeft());
		//Note. can not catch evaluate exception here
		rng.setEditText(formula);
	}

	public void addEventListener(String evtnm, EventListener listener) {
		spreadsheet.addEventListener(evtnm, listener);
	}
	
	public boolean removeEventListener(String evtnm, EventListener listener) {
		return spreadsheet.removeEventListener(evtnm, listener);
	}

	public String getCurrentCellPosition() {
		return (String)spreadsheet.getColumntitle(spreadsheet.getSelection().getLeft()) +
			(String)spreadsheet.getRowtitle(spreadsheet.getSelection().getTop());
	}

	public void setDataFormat(String format) {
		Utils.setDataFormat(spreadsheet.getSelectedSheet(), 
				spreadsheet.getSelection(), format);
	}

	public void save() {
		FileHelper.saveSpreadsheet(spreadsheet);
		//Note. if book come from newBook(), it doesn't store book inside desktop
		storeBookInDesktop(spreadsheet);
	}

	public ByteArrayOutputStream exportToExcel() {
		Book wb = spreadsheet.getBook();
	    Exporter c = Exporters.getExporter("excel");
	    ByteArrayOutputStream out = new ByteArrayOutputStream();
	    c.export(wb, out);
		return out;
	}

	public String getBookName() {
		Book b = spreadsheet.getBook();
		return b != null ? b.getBookName() : null;
	}

	public void close() {
		unsubscribeBookListeners();
		removeBookFromDesktopIfNeeded();
		spreadsheet.setSrcName(null);
		spreadsheet.setBook(null);
	}

	public void addBookEventListener(EventListener listener) {
		Book book = spreadsheet.getBook();
		bookListeners.put(listener, book != null ? Boolean.TRUE : Boolean.FALSE);
		if (book != null)
			book.subscribe(listener);
	}
	
	public void removeBookEventListener(EventListener listener) {
		bookListeners.remove(listener);
		Book book = spreadsheet.getBook();
		if (book != null)
			book.subscribe(listener);
	}

	/**
	 * Subscribe all book event listener when {@link Book} changed
	 */
	private void resubscribeBookListeners() {
		Book book = spreadsheet.getBook();
		if (book == null)
			return;
		for (EventListener listener : bookListeners.keySet()) {
			if (!bookListeners.get(listener)) {
				book.subscribe(listener);
				bookListeners.put(listener, Boolean.TRUE);
			}
		}
	}
	
	/**
	 * Unsubscribe all book event listener before {@link Book} change
	 */
	private void unsubscribeBookListeners() {
		Book book = spreadsheet.getBook();
		if (book == null)
			return;
		for (EventListener listener : bookListeners.keySet()) {
			boolean subscribed = bookListeners.get(listener);
			if (subscribed) {
				book.unsubscribe(listener);
				bookListeners.put(listener, Boolean.FALSE);
			}
		}
	}
	
	/**
	 * Open new spreadsheet
	 */
	public void newBook() {
		unsubscribeBookListeners();
		removeBookFromDesktopIfNeeded();
		FileHelper.openNewSpreadsheet(spreadsheet);
		//Note: new a empty book doesn't share content, no need to store inside desktop
		resubscribeBookListeners();
	}

	public void setBookSrc(String src) {
		unsubscribeBookListeners();
		final Book targetBook = getBookFromDesktop(src);
		removeBookFromDesktopIfNeeded();
		if (targetBook != null) {
			spreadsheet.setBook(targetBook);
			spreadsheet.setSrcName(src);
		}
		else {
			if (!FileHelper.openSrc(src, spreadsheet)) {
				spreadsheet.setSrc(src);
			}
		}
		storeBookInDesktop(spreadsheet);
		resubscribeBookListeners();
	}
	
	public void openBook(SpreadSheetMetaInfo info) {
		unsubscribeBookListeners();
		final Book targetBook = getBookFromDesktop(info.getSrc());
		removeBookFromDesktopIfNeeded();
		if (targetBook != null) {
			spreadsheet.setBook(targetBook);
			spreadsheet.setSrcName(info.getSrc());
		} else {
			FileHelper.openSpreadsheet(spreadsheet, info);
		}
		storeBookInDesktop(spreadsheet);
		resubscribeBookListeners();
	}
	
	/**
	 * Returns {@link #Book} from desktop scope
	 * <p> Search all books inside desktop by {@link Spreadsheet#getSrc}
	 * @return
	 */
	private Book getBookFromDesktop(String src) {
		if (src == null)
			return null;
		
		HashMap<Book, LinkedHashSet<Spreadsheet>> books = getDesktopBooks();
		final String srcBookName = FileHelper.removeFolderPath(src);
		for (Book book : books.keySet()) {
			if (srcBookName.equals(book.getBookName()))
				return book;
		}
		return null;
	}
	
	/**
	 * Store {@link #Book} and relative {@link #Spreadsheet} inside desktop
	 * @param spreadsheet
	 */
	private void storeBookInDesktop(Spreadsheet spreadsheet) {
		Book book = spreadsheet.getBook();
		HashMap<Book, LinkedHashSet<Spreadsheet>> books = getDesktopBooks();
		LinkedHashSet<Spreadsheet> ss = books.get(book);
		if (ss == null) {
			books.put(book, ss = new LinkedHashSet<Spreadsheet>());
		}
		ss.add(spreadsheet);
	}
	
	/**
	 * Remove {@link #Book} from desktop when others spreadsheet doesn't reference to the book
	 */
	private void removeBookFromDesktopIfNeeded() {
		Book book = spreadsheet.getBook();
		if (book == null)
			return;
		
		boolean hasSpreadsheetRef = false;
		HashMap<Book, LinkedHashSet<Spreadsheet>> books = getDesktopBooks();
		if (!books.containsKey(book))
			return;

		LinkedHashSet<Spreadsheet> ss = books.get(book);
		for (Spreadsheet s : ss) {
			Book b = s.getBook();
			if (!s.equals(spreadsheet) && b != null && b.equals(book)) {
				hasSpreadsheetRef = true;
				break;
			}
		}
		if (!hasSpreadsheetRef) {
			books.remove(book);
		}
	}
	
	/**
	 * Returns all books that store inside desktop
	 * 
	 * <p> each book can be set to multiple spreadsheet; 
	 * if a book contains multiple spreadsheet, means the book 
	 * share content cross multiple spreadsheet. for example, 
	 * if book content changed, each spreadsheet will update content.
	 * 
	 * @return
	 */
	private static HashMap<Book, LinkedHashSet<Spreadsheet>> getDesktopBooks() {
		Desktop desktop = Executions.getCurrent().getDesktop();
		HashMap<Book, LinkedHashSet<Spreadsheet>> storer = (HashMap<Book, LinkedHashSet<Spreadsheet>>) desktop.getAttribute(KEY_DESKTOP_BOOKS);
		if (storer == null)
			desktop.setAttribute(KEY_DESKTOP_BOOKS, storer = new HashMap<Book, LinkedHashSet<Spreadsheet>>());
		return storer;
	}

	public boolean hasBook() {
		return spreadsheet.getBook() != null;
	}

	public String getSrc() {
		return spreadsheet.getSrc();
	}

	public void setSrcName(String src) {
		unsubscribeBookListeners();
		spreadsheet.setSrcName(src);
		resubscribeBookListeners();
	}

	public boolean hasFileExtentionName() {
		return FileHelper.isSupportedSpreadSheetExtention(spreadsheet.getSrc());
	}

	public void setColumnWidthInPx(int width) {
		Rect rect = spreadsheet.getSelection();
		final int char256 = Utils.pxToFileChar256(width, ((Book)spreadsheet.getSelectedSheet().getWorkbook()).getDefaultCharWidth());
		Ranges.range(spreadsheet.getSelectedSheet(), 0, rect.getLeft(), 0, rect.getRight()).getColumns().setColumnWidth(char256);
	}

	public void setRowHeightInPx(int height) {
		Rect rect = spreadsheet.getSelection();
		int point = Utils.pxToPoint(height);
		Ranges.range(spreadsheet.getSelectedSheet(), rect.getTop(), 0, rect.getBottom(), 0).getRows().setRowHeight(point);
	}

	public int getDefaultCharWidth() {
		return ((Book)spreadsheet.getSelectedSheet().getWorkbook()).getDefaultCharWidth();
	}
	
	public List<String> getSheetNames() {
		final Book book = spreadsheet.getBook();
		List<String> names = new ArrayList<String>(book.getNumberOfSheets());
		for (int i = 0; i < book.getNumberOfSheets(); i++) {
			names.add(book.getSheetAt(i).getSheetName());
		}
		return Collections.unmodifiableList(names);
	}

	public int shiftSheetLeft() {
		return SheetHelper.shiftSheetLeft(spreadsheet);
	}

	public int shiftSheetRight() {
		return SheetHelper.shiftSheetRight(spreadsheet);
	}

	public int deleteSheet() {
		return SheetHelper.deleteSheet(spreadsheet);
	}
}