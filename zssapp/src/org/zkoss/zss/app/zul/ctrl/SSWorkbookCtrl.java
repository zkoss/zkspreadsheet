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
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.List;

import org.zkoss.image.AImage;
import org.zkoss.poi.ss.formula.eval.NotImplementedException;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.ClientAnchor;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.poi.ss.usermodel.charts.CategoryData;
import org.zkoss.poi.ss.usermodel.charts.ChartData;
import org.zkoss.poi.ss.usermodel.charts.ChartDataSource;
import org.zkoss.poi.ss.usermodel.charts.ChartGrouping;
import org.zkoss.poi.ss.usermodel.charts.ChartTextSource;
import org.zkoss.poi.ss.usermodel.charts.ChartType;
import org.zkoss.poi.ss.usermodel.charts.DataSources;
import org.zkoss.poi.ss.usermodel.charts.LegendPosition;
import org.zkoss.poi.ss.usermodel.charts.XYData;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.poi.xssf.usermodel.XSSFClientAnchor;
import org.zkoss.poi.xssf.usermodel.charts.XSSFArea3DChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFAreaChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFBar3DChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFBarChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFColumn3DChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFColumnChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFDoughnutChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFLine3DChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFLineChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFPie3DChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFPieChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFScatChartData;
import org.zkoss.util.resource.Labels;
import org.zkoss.zk.ui.Desktop;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.util.Clients;
import org.zkoss.zss.app.cell.CellHelper;
import org.zkoss.zss.app.file.FileHelper;
import org.zkoss.zss.app.file.SpreadSheetMetaInfo;
import org.zkoss.zss.app.sheet.SheetHelper;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XExporter;
import org.zkoss.zss.model.sys.XExporters;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.XRanges;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.ui.Position;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.CellEvent;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.impl.HeaderPositionHelper;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zss.ui.sys.SpreadsheetCtrl;

/**
 * 
 * @author Sam
 *
 */
public class SSWorkbookCtrl implements WorkbookCtrl {

	private Spreadsheet spreadsheet;
	
	private String lastSheetName = null;

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

	public void insertColumnLeft() {
		if (spreadsheet.getSelection() == null)
			return;
		
		Rect rect = spreadsheet.getSelection();
		CellHelper.shiftEntireColumnRight(spreadsheet.getSelectedXSheet(), 
				rect.getLeft(), rect.getRight());
	}

	public void deleteColumn() {
		Rect rect = spreadsheet.getSelection();
		if (rect == null)
			return;
		
		CellHelper.shiftEntireColumnLeft(spreadsheet.getSelectedXSheet(), 
				rect.getLeft(), rect.getRight());
	}

	public void insertRowAbove() {
		CellHelper.shiftEntireRowDown(spreadsheet.getSelectedXSheet(), 
				spreadsheet.getSelection().getTop(), 
				spreadsheet.getSelection().getBottom());
	}

	public void deleteRow() {
		Rect rect = spreadsheet.getSelection();
		CellHelper.shiftEntireRowUp(spreadsheet.getSelectedXSheet(), rect.getTop(), rect.getBottom());
	}

	public void setSelectedSheet(String name) {
		XSheet lastsheet = spreadsheet.getSelectedXSheet();
		
		Integer lastIdx = Integer.valueOf(spreadsheet.getXBook().getSheetIndex(lastsheet));
		
		spreadsheet.setSelectedSheet(name);		
		//TODO: handle the copy/cut highlight
//		final Worksheet sheet = EditHelper.getSourceSheet(spreadsheet);
//		if (sheet != null) {
//			if (sheet.equals(spreadsheet.getSelectedSheet())) {
//				spreadsheet.setHighlight(EditHelper.getSourceRange(spreadsheet));
//			} else {
//				spreadsheet.setHighlight(null);
//			}
//		}
		
		lastSheetName = name;
		spreadsheet.focus(); //move focus in to spreadsheet
	}

	public void hide(boolean hide) {
		Rect rect = spreadsheet.getSelection();
		if (rect == null)
			return;
		
		XRanges.range(spreadsheet.getSelectedXSheet(), 
				rect.getTop(), rect.getLeft(), 
				rect.getBottom(), rect.getRight()).setHidden(hide);
	}

	public void addImage(int row, int col, AImage image) {
		if (WebApps.getFeature("pe")) {
			XRanges.range(spreadsheet.getSelectedXSheet()).addPicture(getClientCenterAnchor(row, col, image.getWidth(), image.getHeight()), image.getByteData(), getImageFormat(image));
		}
	}
	
	private int getImageFormat(AImage image) {
		String format = image.getFormat();
		if ("dib".equalsIgnoreCase(format)) {
			return Workbook.PICTURE_TYPE_DIB;
		} else if ("emf".equalsIgnoreCase(format)) {
			return Workbook.PICTURE_TYPE_EMF;
		} else if ("wmf".equalsIgnoreCase(format)) {
			return Workbook.PICTURE_TYPE_WMF;
		} else if ("jpeg".equalsIgnoreCase(format)) {
			return Workbook.PICTURE_TYPE_JPEG;
		} else if ("pict".equalsIgnoreCase(format)) {
			return Workbook.PICTURE_TYPE_PICT;
		} else if ("png".equalsIgnoreCase(format)) {
			return Workbook.PICTURE_TYPE_PNG;
		}
		throw new UiException("Unsupported format: " + format);
	}

	public void insertSheet() {
		int sheetCount = spreadsheet.getXBook().getNumberOfSheets();
		XRanges.range(spreadsheet.getSelectedXSheet()).createSheet("Sheet " + (sheetCount + 1));
	}

	public void reGainFocus() {
		Clients.evalJavaScript("zk.Widget.$('" + spreadsheet.getUuid() + "').focus(false);");
	}

	public void renameSelectedSheet(String name) {
		XRanges.range(spreadsheet.getSelectedXSheet()).setSheetName(name);
	}

	public void shiftCell(int direction) {
		XSheet sheet = spreadsheet.getSelectedXSheet();
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
			CellHelper.sortDescending(spreadsheet.getSelectedXSheet(), SheetHelper.getSpreadsheetMaxSelection(spreadsheet));
		else
			CellHelper.sortAscending(spreadsheet.getSelectedXSheet(), SheetHelper.getSpreadsheetMaxSelection(spreadsheet));
	}

	public void setColumnFreeze(int columnFreeze) {
		spreadsheet.setColumnfreeze(columnFreeze);
	}

	public void setRowFreeze(int rowfreeze) {
		spreadsheet.setRowfreeze(rowfreeze);
	}

	public void insertFormula(int rowIdx, int colIdx, String formula) {
		XRange rng = XRanges.range(spreadsheet.getSelectedXSheet(), rowIdx, colIdx);
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
		Utils.setDataFormat(spreadsheet.getSelectedXSheet(), 
				spreadsheet.getSelection(), format);
	}

	public void save() {
		FileHelper.saveSpreadsheet(spreadsheet);
		//Note. if book come from newBook(), it doesn't store book inside desktop
		storeBookInDesktop(spreadsheet);
	}

	public ByteArrayOutputStream exportToExcel() {
		XBook wb = spreadsheet.getXBook();
	    XExporter c = XExporters.getExporter("excel");
	    ByteArrayOutputStream out = new ByteArrayOutputStream();
	    c.export(wb, out);
		return out;
	}

	public String getBookName() {
		XBook b = spreadsheet.getXBook();
		return b != null ? b.getBookName() : null;
	}

	public void close() {
		unsubscribeBookListeners();
		removeBookFromDesktopIfNeeded();
		spreadsheet.setSrcName(null);
		spreadsheet.setBook(null);
		
		spreadsheet.getUserActionHandler().toggleActionOnBookClosed();
	}

	public void addBookEventListener(EventListener listener) {
		XBook book = spreadsheet.getXBook();
		bookListeners.put(listener, book != null ? Boolean.TRUE : Boolean.FALSE);
		if (book != null)
			book.subscribe(listener);
	}
	
	public void removeBookEventListener(EventListener listener) {
		bookListeners.remove(listener);
		XBook book = spreadsheet.getXBook();
		if (book != null)
			book.subscribe(listener);
	}

	/**
	 * Subscribe all book event listener when {@link XBook} changed
	 */
	private void resubscribeBookListeners() {
		XBook book = spreadsheet.getXBook();
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
	 * Unsubscribe all book event listener before {@link XBook} change
	 */
	private void unsubscribeBookListeners() {
		XBook book = spreadsheet.getXBook();
		if (book == null) {
			for (EventListener listener : bookListeners.keySet()) {
				bookListeners.put(listener, Boolean.FALSE);
			}
			return;
		}
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
		final XBook targetBook = getBookFromDesktop(src);
		removeBookFromDesktopIfNeeded();
		if (targetBook != null) {
			spreadsheet.setXBook(targetBook);
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
	
	public void setBook(XBook book) {
		unsubscribeBookListeners();
		removeBookFromDesktopIfNeeded();
		spreadsheet.setXBook(book);
		storeBookInDesktop(spreadsheet);
		resubscribeBookListeners();
	}
	
	public void openBook(SpreadSheetMetaInfo info) {
		unsubscribeBookListeners();
		final XBook targetBook = getBookFromDesktop(info.getSrc());
		removeBookFromDesktopIfNeeded();
		if (targetBook != null) {
			spreadsheet.setXBook(targetBook);
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
	private XBook getBookFromDesktop(String src) {
		if (src == null)
			return null;
		
		HashMap<XBook, LinkedHashSet<Spreadsheet>> books = getDesktopBooks();
		final String srcBookName = FileHelper.removeFolderPath(src);
		for (XBook book : books.keySet()) {
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
		XBook book = spreadsheet.getXBook();
		HashMap<XBook, LinkedHashSet<Spreadsheet>> books = getDesktopBooks();
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
		XBook book = spreadsheet.getXBook();
		if (book == null)
			return;
		
		boolean hasSpreadsheetRef = false;
		HashMap<XBook, LinkedHashSet<Spreadsheet>> books = getDesktopBooks();
		if (!books.containsKey(book))
			return;

		LinkedHashSet<Spreadsheet> ss = books.get(book);
		for (Spreadsheet s : ss) {
			XBook b = s.getXBook();
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
	private static HashMap<XBook, LinkedHashSet<Spreadsheet>> getDesktopBooks() {
		Desktop desktop = Executions.getCurrent().getDesktop();
		HashMap<XBook, LinkedHashSet<Spreadsheet>> storer = (HashMap<XBook, LinkedHashSet<Spreadsheet>>) desktop.getAttribute(KEY_DESKTOP_BOOKS);
		if (storer == null)
			desktop.setAttribute(KEY_DESKTOP_BOOKS, storer = new HashMap<XBook, LinkedHashSet<Spreadsheet>>());
		return storer;
	}

	public boolean hasBook() {
		return spreadsheet.getXBook() != null;
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

	public void setColumnWidthInPx(int width, Rect selection) {
		final int char256 = Utils.pxToFileChar256(width, ((XBook)spreadsheet.getSelectedXSheet().getWorkbook()).getDefaultCharWidth());
		XRanges.range(spreadsheet.getSelectedXSheet(), 0, selection.getLeft(), 0, selection.getRight()).getColumns().setColumnWidth(char256);
	}

	public void setRowHeightInPx(int height, Rect selection) {
		int point = Utils.pxToPoint(height);
		XRanges
		.range(spreadsheet.getSelectedXSheet(), selection.getTop(), 0, selection.getBottom(), 0)
		.getRows()
		.setRowHeight(point);
	}

	public int getDefaultCharWidth() {
		return ((XBook)spreadsheet.getSelectedXSheet().getWorkbook()).getDefaultCharWidth();
	}
	
	public List<String> getSheetNames() {
		final XBook book = spreadsheet.getXBook();
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

	public boolean isSheetProtect() {
		return spreadsheet.getSelectedXSheet().getProtect();
	}
	
	public void protectSheet(String password) {
		XRanges.range(spreadsheet.getSelectedXSheet()).protectSheet(password);
	}

	public XSheet getSelectedSheet() {
		return spreadsheet.getSelectedXSheet();
	}

	public String getReference(int row, int column) {
		return (String)spreadsheet.getColumntitle(column) + spreadsheet.getRowtitle(row);
	}

	public void escapeAndUpdateText(org.zkoss.poi.ss.usermodel.Cell cell, String text) {
		spreadsheet.escapeAndUpdateText(cell, text);
	}

	public void focusTo(int row, int column, boolean fireFocusEvent) {
		spreadsheet.focusTo(row, column);
		if (fireFocusEvent) {
			org.zkoss.zk.ui.event.Events.sendEvent(
					new CellEvent(Events.ON_CELL_FOUCSED, spreadsheet, spreadsheet.getSelectedXSheet(), row, column));
		}
	}

	public void moveEditorFocus(String id, String name, String color, int row, int col) {
		spreadsheet.moveEditorFocus(id, name, color, row, col);
	}

	public void removeEditorFocus(String id) {
		spreadsheet.removeEditorFocus(id);
	}

	public Rect getSelection() {
		return spreadsheet.getSelection();
	}

	public Position getCellFocus() {
		return spreadsheet.getCellFocus();
	}

	public int getMaxcolumns() {
		return spreadsheet.getMaxcolumns();
	}

	public int getMaxrows() {
		return spreadsheet.getMaxrows();
	}

	//TODO: rm this
	public void clearClipbook() {
//		EditHelper.clearCutOrCopy(spreadsheet);
	}

	public void updateText(Cell cell, String text) {
		spreadsheet.updateText(cell, text);
	}
	
	public String getColumnTitle(int col) {
		return Labels.getLabel("column") + " " + spreadsheet.getColumntitle(col);
	}

	public String getRowTitle(int row) {
		return Labels.getLabel("row") + " " + spreadsheet.getRowtitle(row);
	}
	
	public Rect getVisibleRect() {
		SpreadsheetCtrl ctrl = (SpreadsheetCtrl) spreadsheet.getExtraCtrl();
		return ctrl.getVisibleRect();
	}

	public void addChart(int row, int col, ChartType chartType) {
		XSheet sheet = spreadsheet.getSelectedXSheet();
		
		Rect selection = spreadsheet.getSelection();
		XRanges.range(sheet).addChart(getClientCenterAnchor(row, col, 600, 300), newChartData(chartType, selection), chartType, ChartGrouping.STANDARD, LegendPosition.RIGHT);
	}
	
	private ChartData newChartData(ChartType chartType, Rect selection) {
		ChartData data = null;
		
		switch (chartType) {
		case Area3D:
			data = fillCategoryData(new XSSFArea3DChartData());
			break;
		case Area:
			data = fillCategoryData(new XSSFAreaChartData());
			break;
		case Bar3D:
			data = fillCategoryData(new XSSFBar3DChartData());
			//((XSSFBar3DChartData) data).setGrouping(ChartGrouping.STANDARD);
			break;
		case Column3D:
			data = fillCategoryData(new XSSFColumn3DChartData());
			//((XSSFBar3DChartData) data).setGrouping(ChartGrouping.STANDARD);
			break;
		case Bar:
			data = fillCategoryData(new XSSFBarChartData());
			//((XSSFBarChartData) data).setGrouping(ChartGrouping.STANDARD);
			break;
		case Column:
			data = fillCategoryData(new XSSFColumnChartData());
			//((XSSFBarChartData) data).setGrouping(ChartGrouping.STANDARD);
			break;
		case Bubble:
			throw new UnsupportedOperationException();
		case Doughnut:
			data = fillCategoryData(new XSSFDoughnutChartData());
			break;
		case Line3D:
			data = fillCategoryData(new XSSFLine3DChartData());
			break;
		case Line:
			data = fillCategoryData(new XSSFLineChartData());
			break;
		case Pie3D:
			data = fillCategoryData(new XSSFPie3DChartData());
			break;
		case OfPie:
//			break;
			throw new UnsupportedOperationException();
		case Pie:
			data = fillCategoryData(new XSSFPieChartData());
			break;
		case Radar:
			throw new NotImplementedException("Radar data not impl");
		case Scatter:
			data = fillXYData(new XSSFScatChartData());
			break;
		case Stock:
//			data = fillCategoryData(new XSSFStockChartData());
//			break;
			throw new UnsupportedOperationException();
		case Surface3D:
//			break;
			throw new UnsupportedOperationException();
		case Surface:
//			break;
			throw new UnsupportedOperationException();
		}
		return data;
	}
	
	private ChartData fillXYData(XYData data) {
		final Rect selection = spreadsheet.getSelection();
		final XSheet sheet = spreadsheet.getSelectedXSheet();
		
		Rect rect = getChartDataRange(selection, sheet);
		int colIdx = rect.getLeft();
		int rowIdx = rect.getTop();
		
		ChartDataSource<Number> horValues = null;
		ArrayList<ChartTextSource> titles = new ArrayList<ChartTextSource>();
		ArrayList<ChartDataSource<Number>> values = new ArrayList<ChartDataSource<Number>>();
		
		int colWidth = selection.getRight() - colIdx;
		int rowHeight = selection.getBottom() - rowIdx;
		if (rowHeight > colWidth) {
			//find horizontal value, at least 1 column
			if (colIdx < selection.getRight()) {
				int lCol = selection.getLeft();
				int rCol = lCol;
				if (rCol < colIdx) {
					rCol = colIdx - 1;
				} else {
					colIdx += 1;
				}
				String startCell = spreadsheet.getColumntitle(lCol) + spreadsheet.getRowtitle(rowIdx);
				String endCell = spreadsheet.getColumntitle(rCol) + spreadsheet.getRowtitle(selection.getBottom());
				horValues = DataSources.fromNumericCellRange(sheet, CellRangeAddress.valueOf(startCell + ":" + endCell));
			}
			//find values
			int i = 1;
			for (int c = colIdx; c <= selection.getRight(); c++) {
				//find title
				String title = null;
				int row = rowIdx - 1;
				if (row >= selection.getTop()) {
					title = "" + XRanges.range(sheet, selection.getTop(), c, row, c).getText().toString();
				}
				titles.add(title == null ? null : DataSources.fromString(title));
				
				String startCell = spreadsheet.getColumntitle(c) + spreadsheet.getRowtitle(rowIdx);
				String endCell = spreadsheet.getColumntitle(c) + spreadsheet.getRowtitle(selection.getBottom());
				values.add(DataSources.fromNumericCellRange(sheet, CellRangeAddress.valueOf(startCell + ":" + endCell)));
			}
		} else {
			//find horizontal value, at least 1 row
			if (rowIdx < selection.getBottom()) {
				int tRow = selection.getTop();
				int bRow = tRow;
				if (bRow < rowIdx) {
					bRow = rowIdx - 1;
				} else {
					rowIdx += 1;
				}
				String startCell = spreadsheet.getColumntitle(colIdx) + spreadsheet.getRowtitle(tRow);
				String endCell = spreadsheet.getColumntitle(selection.getRight()) + spreadsheet.getRowtitle(tRow);
				horValues = DataSources.fromNumericCellRange(sheet, CellRangeAddress.valueOf(startCell + ":" + endCell));
			}
			//find values
			int i = 1;
			for (int r = rowIdx; r <= selection.getBottom(); r++) {
				//find title
				String title = null;
				int col = colIdx - 1;
				if (col >= selection.getLeft()) {
					title = "" + XRanges.range(sheet, r, selection.getLeft(), r, col).getText().toString();
				}
				titles.add(title == null ? null : DataSources.fromString(title));
				
				String startCell = spreadsheet.getColumntitle(colIdx) + spreadsheet.getRowtitle(r);
				String endCell = spreadsheet.getColumntitle(selection.getRight()) + spreadsheet.getRowtitle(r);
				values.add(DataSources.fromNumericCellRange(sheet, CellRangeAddress.valueOf(startCell + ":" + endCell)));
			}
		}
		
		for (int i = 0; i < values.size(); i++) {
			data.addSerie(titles.get(i), horValues, values.get(i));
		}
		return data;
	}

	private Rect getChartDataRange(Rect selection, XSheet sheet) {
		//assume can't find number cell, use last cell as value
		int colIdx = selection.getLeft();
		int rowIdx = -1;
		for (int r = selection.getBottom(); r >= selection.getTop(); r--) {
			Row row = sheet.getRow(r);
			int rCol = colIdx;
			for (int c = selection.getRight(); c >= rCol; c--) {
				if (isQualifiedCell(row.getCell(c))) {
					colIdx = c;
					rowIdx = r;
				} else {
					break;
				}
			}
		}
		if (rowIdx == -1) { //can not find number cell, use last cell as chart's value
			rowIdx = selection.getBottom();
			colIdx = selection.getRight();
		}
		return new Rect(colIdx, rowIdx, selection.getRight(), selection.getBottom());
	}
	
	private CategoryData fillCategoryData(CategoryData data) {
		final Rect selection = spreadsheet.getSelection();
		final XSheet sheet = spreadsheet.getSelectedXSheet();
		
		Rect rect = getChartDataRange(selection, sheet);
		int colIdx = rect.getLeft();
		int rowIdx = rect.getTop();
		
		ChartDataSource<String> cats = null;
		ArrayList<ChartTextSource> titles = new ArrayList<ChartTextSource>();
		ArrayList<ChartDataSource<Number>> vals = new ArrayList<ChartDataSource<Number>>();
		
		int colWidth = selection.getRight() - colIdx;
		int rowHeight = selection.getBottom() - rowIdx;
		if (rowHeight > colWidth) { //catalog by row, value by column
			//find catalog
			int col = colIdx - 1;
			if (col >= selection.getLeft()) {
				String startCell = spreadsheet.getColumntitle(selection.getLeft()) + spreadsheet.getRowtitle(rowIdx);
				String endCell = spreadsheet.getColumntitle(col) + spreadsheet.getRowtitle(selection.getBottom());
				cats = DataSources.fromStringCellRange(sheet, CellRangeAddress.valueOf(startCell + ":" + endCell));
			}
			//find value, by column
			int i = 1;
			for (int c = colIdx; c <= selection.getRight(); c++) {
				//find title
				String title = null;
				int row = rowIdx - 1;
				if (row >= selection.getTop()) {
					title = "" + XRanges.range(sheet, selection.getTop(), c, row, c).getText().toString();
				}
				titles.add(title == null ? null : DataSources.fromString(title));
				
				String startCell = spreadsheet.getColumntitle(c) + spreadsheet.getRowtitle(rowIdx);
				String endCell = spreadsheet.getColumntitle(c) + spreadsheet.getRowtitle(selection.getBottom());
				ChartDataSource<Number> val = DataSources.fromNumericCellRange(sheet, CellRangeAddress.valueOf(startCell + ":" + endCell));
				vals.add(val);
			}
		} else { //catalog by column, value by row
			//find catalog
			int row = rowIdx - 1;
			if (row >= selection.getTop()) {
				String startCell = spreadsheet.getColumntitle(colIdx) + spreadsheet.getRowtitle(row);
				String endCell = spreadsheet.getColumntitle(selection.getRight()) + spreadsheet.getRowtitle(row);
				cats = DataSources.fromStringCellRange(sheet, CellRangeAddress.valueOf(startCell + ":" + endCell));
			}
			
			//find value
			int i = 1;
			for (int r = rowIdx; r <= selection.getBottom(); r++) {
				//find title
				String title = null;
				int col = colIdx - 1;
				if (col >= selection.getLeft()) {
					title = "" + XRanges.range(sheet, r, selection.getLeft(), r, col).getText().toString();
				}
				titles.add(title == null ? null : DataSources.fromString(title));
				
				String startCell = spreadsheet.getColumntitle(colIdx) + spreadsheet.getRowtitle(r);
				String endCell = spreadsheet.getColumntitle(selection.getRight()) + spreadsheet.getRowtitle(r);
				ChartDataSource<Number> val = DataSources.fromNumericCellRange(sheet, CellRangeAddress.valueOf(startCell + ":" + endCell));
				vals.add(val);
			}
		}
		
		for (int i = 0; i < vals.size(); i++) {
			data.addSerie(titles.get(i), cats, vals.get(i));
		}
		return data;
	}
	
	private boolean isQualifiedCell(Cell cell) {
		if (cell == null)
			return true;
		int cellType = cell.getCellType();
		return cellType == Cell.CELL_TYPE_NUMERIC || 
			cellType == Cell.CELL_TYPE_FORMULA || 
			cellType == Cell.CELL_TYPE_BLANK;
	}
	
	private ClientAnchor getClientCenterAnchor(int row, int col, int widgetWidth, int widgetHeight) {
		HeaderPositionHelper rowSizeHelper = (HeaderPositionHelper) spreadsheet.getAttribute("_rowCellSize");
		HeaderPositionHelper colSizeHelper = (HeaderPositionHelper) spreadsheet.getAttribute("_colCellSize");
		
		int lCol = col;
		int tRow = row;
		int rCol = lCol;
		int bRow = tRow;
		int offsetWidth = 0;
		int offsetHeight = 0;
		for (int r = tRow; r < spreadsheet.getMaxrows(); r++) {
			int cellHeight = rowSizeHelper.getSize(r);
			widgetHeight -= cellHeight;
			if (widgetHeight <= 0) {
				bRow = r;
				if (widgetHeight < 0) {
					offsetHeight = cellHeight - Math.abs(widgetHeight);
				}
				break;
			}
		}
		for (int c = lCol; c < spreadsheet.getMaxcolumns(); c++) {
			int cellWidth = colSizeHelper.getSize(c);
			widgetWidth -= cellWidth;
			if (widgetWidth <= 0) {
				rCol = c;
				if (widgetWidth < 0) {
					offsetWidth = cellWidth - Math.abs(widgetWidth);
				}
				break;
			}
		}
		ClientAnchor anchor = new XSSFClientAnchor(0, 0, pxToEmu(offsetWidth), pxToEmu(offsetHeight), lCol, tRow, rCol, bRow);
		return anchor;
	}
	
	public boolean setEditTextWithValidation(XSheet sheet, int row, int col, String txt, EventListener okCallback) {
		return Utils.setEditTextWithValidation(spreadsheet, sheet, row, col, txt, okCallback);
	}
	
	/** convert pixel to EMU */
	public static int pxToEmu(int px) {
		return (int) Math.round(((double)px) * 72 * 20 * 635 / 96); //assume 96dpi
	}
}