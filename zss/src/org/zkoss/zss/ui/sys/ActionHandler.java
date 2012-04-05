/* ToolbarActionHandler.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Feb 24, 2012 2:08:34 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ui.sys;

import java.util.ArrayList;
import java.util.Map;

import org.zkoss.image.AImage;
import org.zkoss.lang.Objects;
import org.zkoss.lang.Strings;
import org.zkoss.poi.ss.usermodel.AutoFilter;
import org.zkoss.poi.ss.usermodel.BorderStyle;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.ClientAnchor;
import org.zkoss.poi.ss.usermodel.Font;
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
import org.zkoss.util.media.Media;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.UploadEvent;
import org.zkoss.zss.engine.event.SSDataEvent;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.ui.Action;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.event.KeyEvent;
import org.zkoss.zss.ui.impl.CellVisitor;
import org.zkoss.zss.ui.impl.CellVisitorContext;
import org.zkoss.zss.ui.impl.HeaderPositionHelper;
import org.zkoss.zss.ui.impl.MergeMatrixHelper;
import org.zkoss.zss.ui.impl.MergedRect;
import org.zkoss.zss.ui.impl.Upload;
import org.zkoss.zss.ui.impl.Uploader;
import org.zkoss.zss.ui.impl.Utils;

/**
 * @author sam
 *
 */
public abstract class ActionHandler {
	
	protected Spreadsheet _spreadsheet;
	protected Uploader _insertPicture;
	protected Rect _insertPictureSelection;
	protected Book _book;
	protected Clipboard _clipboard;
	protected Action[] disabledActionOnBookClose = new Action[]{
			Action.SAVE_BOOK,
			Action.EXPORT_PDF, 
			Action.PASTE,
			Action.CUT,
			Action.COPY,
			Action.FONT_FAMILY,
			Action.FONT_SIZE,
			Action.FONT_BOLD,
			Action.FONT_ITALIC,
			Action.FONT_UNDERLINE,
			Action.FONT_STRIKE,
			Action.BORDER,
			Action.FONT_COLOR,
			Action.FILL_COLOR,
			Action.VERTICAL_ALIGN,
			Action.HORIZONTAL_ALIGN,
			Action.WRAP_TEXT,
			Action.MERGE_AND_CENTER,
			Action.INSERT,
			Action.DELETE,
			Action.CLEAR,
			Action.SORT_AND_FILTER,
			Action.PROTECT_SHEET,
			Action.GRIDLINES,
			Action.INSERT_PICTURE,
			Action.COLUMN_CHART,
			Action.LINE_CHART,
			Action.PIE_CHART,
			Action.BAR_CHART,
			Action.AREA_CHART,
			Action.SCATTER_CHART,
			Action.OTHER_CHART,
			Action.HYPERLINK
	};
	protected EventListener _bookListener = new EventListener() {

		@Override
		public void onEvent(Event event) throws Exception {
			SSDataEvent evt = (SSDataEvent)event;
			if (evt.getName().equals(SSDataEvent.ON_BTN_CHANGE)) {
				syncAutoFilter();
			}
		}
	};
	
	public ActionHandler() {}
	public ActionHandler(Spreadsheet spreadsheet) {
		_spreadsheet = spreadsheet;
		init();
	}
	
	public void dispatch(String toolbarAction, Map data) {
		if (Action.HOME_PANEL.toString().equals(toolbarAction)) {
			doHomePanel();
		} else if (Action.INSERT_PANEL.equals(toolbarAction)) {
			doInsertPanel();
		} else if (Action.FORMULA_PANEL.equals(toolbarAction)) {
			doFormulaPanel();
		} else if (Action.NEW_BOOK.equals(toolbarAction)) {
			doNewBook();
		} else if (Action.SAVE_BOOK.equals(toolbarAction)) {
			doSaveBook();
		} else if (Action.EXPORT_PDF.equals(toolbarAction)) {
			doExportPDF(getSelection(data));
		} else if (Action.PASTE.equals(toolbarAction)) {
			doPaste(getSelection(data));
		} else if (Action.PASTE_FORMULA.equals(toolbarAction)) {
			doPasteFormula(getSelection(data));
		} else if (Action.PASTE_VALUE.equals(toolbarAction)) {
			doPasteValue(getSelection(data));
		} else if (Action.PASTE_ALL_EXPECT_BORDERS.equals(toolbarAction)) {
			doPasteAllExceptBorder(getSelection(data));
		} else if (Action.PASTE_TRANSPOSE.equals(toolbarAction)) {
			doPasteTranspose(getSelection(data));
		} else if (Action.PASTE_SPECIAL.equals(toolbarAction)) {
			doPasteSpecial(getSelection(data));
		} else if (Action.CUT.equals(toolbarAction)) {
			doCut(getSelection(data));	
		} else if (Action.COPY.equals(toolbarAction)) {
			doCopy(getSelection(data));
		} else if (Action.FONT_FAMILY.equals(toolbarAction)) {
			doFontFamily((String)data.get("name"), getSelection(data));
		} else if (Action.FONT_SIZE.equals(toolbarAction)) {
			Integer fontSize = Integer.parseInt((String)data.get("size"));
			doFontSize(fontSize, getSelection(data));
		} else if (Action.FONT_BOLD.equals(toolbarAction)) {
			doFontBold(getSelection(data));
		} else if (Action.FONT_ITALIC.equals(toolbarAction)) {
			doFontItalic(getSelection(data));
		} else if (Action.FONT_UNDERLINE.equals(toolbarAction)) {
			doFontUnderline(getSelection(data));
		} else if (Action.FONT_STRIKE.equals(toolbarAction)) {
			doFontStrikeout(getSelection(data));
		} else if (Action.BORDER.equals(toolbarAction)) {
			String color = (String)data.get("color");
			if (Strings.isEmpty(color)) {//CE version won't provide color
				color = getBorderColor();
			}
			doBorder(color, getSelection(data));
		} else if (Action.BORDER_BOTTOM.equals(toolbarAction)) {
			String color = (String)data.get("color");
			if (Strings.isEmpty(color)) {
				color = getBorderColor();
			}
			doBorderBottom(color, getSelection(data));
		} else if (Action.BORDER_TOP.equals(toolbarAction)) {
			String color = (String)data.get("color");
			if (Strings.isEmpty(color)) {
				color = getBorderColor();
			}
			doBoderTop(color, getSelection(data));
		} else if (Action.BORDER_LEFT.equals(toolbarAction)) {
			String color = (String)data.get("color");
			if (Strings.isEmpty(color)) {
				color = getBorderColor();
			}
			doBorderLeft(color, getSelection(data));
		} else if (Action.BORDER_RIGHT.equals(toolbarAction)) {
			String color = (String)data.get("color");
			if (Strings.isEmpty(color)) {
				color = getBorderColor();
			}
			doBorderRight(color, getSelection(data));
		} else if (Action.BORDER_NO.equals(toolbarAction)) {
			String color = (String)data.get("color");
			if (Strings.isEmpty(color)) {
				color = getBorderColor();
			}
			doBorderNo(color, getSelection(data));
		} else if (Action.BORDER_ALL.equals(toolbarAction)) {
			String color = (String)data.get("color");
			if (Strings.isEmpty(color)) {
				color = getBorderColor();
			}
			doBorderAll(color, getSelection(data));
		} else if (Action.BORDER_OUTSIDE.equals(toolbarAction)) {
			String color = (String)data.get("color");
			if (Strings.isEmpty(color)) {
				color = getBorderColor();
			}
			doBorderOutside(color, getSelection(data));
		} else if (Action.BORDER_INSIDE.equals(toolbarAction)) {
			String color = (String)data.get("color");
			if (Strings.isEmpty(color)) {
				color = getBorderColor();
			}
			doBorderInside(color, getSelection(data));
		} else if (Action.BORDER_INSIDE_HORIZONTAL.equals(toolbarAction)) {
			String color = (String)data.get("color");
			if (Strings.isEmpty(color)) {
				color = getBorderColor();
			}
			doBorderInsideHorizontal(color, getSelection(data));
		} else if (Action.BORDER_INSIDE_VERTICAL.equals(toolbarAction)) {
			String color = (String)data.get("color");
			if (Strings.isEmpty(color)) {
				color = getBorderColor();
			}
			doBorderInsideVertical(color, getSelection(data));
		} else if (Action.FONT_COLOR.equals(toolbarAction)) {
			doFontColor((String)data.get("color"), getSelection(data));
		} else if (Action.FILL_COLOR.equals(toolbarAction)) {
			doFillColor((String)data.get("color"), getSelection(data));
		} else if (Action.VERTICAL_ALIGN_TOP.equals(toolbarAction)) {
			doVerticalAlignTop(getSelection(data));
		} else if (Action.VERTICAL_ALIGN_MIDDLE.equals(toolbarAction)) {
			doVerticalAlignMiddle(getSelection(data));
		} else if (Action.VERTICAL_ALIGN_BOTTOM.equals(toolbarAction)) {
			doVerticalAlignBottom(getSelection(data));
		} else if (Action.HORIZONTAL_ALIGN_LEFT.equals(toolbarAction)) {
			doHorizontalAlignLeft(getSelection(data));
		} else if (Action.HORIZONTAL_ALIGN_CENTER.equals(toolbarAction)) {
			doHorizontalAlignCenter(getSelection(data));
		} else if (Action.HORIZONTAL_ALIGN_RIGHT.equals(toolbarAction)) {
			doHorizontalAlignRight(getSelection(data));
		} else if (Action.WRAP_TEXT.equals(toolbarAction)) {
			doWrapText(getSelection(data));
		} else if (Action.MERGE_AND_CENTER.equals(toolbarAction)) {
			doMergeAndCenter(getSelection(data));
		} else if (Action.MERGE_ACROSS.equals(toolbarAction)) {
			doMergeAcross(getSelection(data));
		} else if (Action.MERGE_CELL.equals(toolbarAction)) {
			doMergeCell(getSelection(data));
		} else if (Action.UNMERGE_CELL.equals(toolbarAction)) {
			doUnmergeCell(getSelection(data));
		} else if (Action.INSERT_SHIFT_CELL_RIGHT.equals(toolbarAction)) {
			doShiftCellRight(getSelection(data));
		} else if (Action.INSERT_SHIFT_CELL_DOWN.equals(toolbarAction)) {
			doShiftCellDown(getSelection(data));
		} else if (Action.INSERT_SHEET_ROW.equals(toolbarAction)) {
			doInsertSheetRow(getSelection(data));
		} else if (Action.INSERT_SHEET_COLUMN.equals(toolbarAction)) {
			doInsertSheetColumn(getSelection(data));
		} else if (Action.DELETE_SHIFT_CELL_LEFT.equals(toolbarAction)) {
			doShiftCellLeft(getSelection(data));
		} else if (Action.DELETE_SHIFT_CELL_UP.equals(toolbarAction)) {
			doShiftCellUp(getSelection(data));
		} else if (Action.DELETE_SHEET_ROW.equals(toolbarAction)) {
			doDeleteSheetRow(getSelection(data));
		} else if (Action.DELETE_SHEET_COLUMN.equals(toolbarAction)) {
			doDeleteSheetColumn(getSelection(data));
		} else if (Action.SORT_ASCENDING.equals(toolbarAction)) {
			doSortAscending(getSelection(data));
		} else if (Action.SORT_DESCENDING.equals(toolbarAction)) {
			doSortDescending(getSelection(data));
		} else if (Action.CUSTOM_SORT.equals(toolbarAction)) {
			doCustomSort(getSelection(data));
		} else if (Action.FILTER.equals(toolbarAction)) {
			doFilter(getSelection(data));
		} else if (Action.CLEAR_FILTER.equals(toolbarAction)) {
			doClearFilter();
		} else if (Action.REAPPLY_FILTER.equals(toolbarAction)) {
			doReapplyFilter();
		} else if (Action.CLEAR_CONTENT.equals(toolbarAction)) {
			doClearContent(getSelection(data));
		} else if (Action.CLEAR_STYLE.equals(toolbarAction)) {
			doClearStyle(getSelection(data));
		} else if (Action.CLEAR_ALL.equals(toolbarAction)) {
			doClearAll(getSelection(data));
		} else if (Action.PROTECT_SHEET.equals(toolbarAction)) {
			doProtectSheet();
		} else if (Action.GRIDLINES.equals(toolbarAction)) {
			doGridlines();
		} else if (Action.COLUMN_CHART.equals(toolbarAction)) {
			doColumnChart(getSelection(data));
		} else if (Action.COLUMN_CHART_3D.equals(toolbarAction)) {
			doColumnChart3D(getSelection(data));
		} else if (Action.LINE_CHART.equals(toolbarAction)) {
			doLineChart(getSelection(data));
		} else if (Action.LINE_CHART_3D.equals(toolbarAction)) {
			doLineChart3D(getSelection(data));
		} else if (Action.PIE_CHART.equals(toolbarAction)) {
			doPieChart(getSelection(data));
		} else if (Action.PIE_CHART_3D.equals(toolbarAction)) {
			doPieChart3D(getSelection(data));
		} else if (Action.BAR_CHART.equals(toolbarAction)) {
			doBarChart(getSelection(data));
		} else if (Action.BAR_CHART_3D.equals(toolbarAction)) {
			doBarChart3D(getSelection(data));
		} else if (Action.AREA_CHART.equals(toolbarAction)) {
			doAreaChart(getSelection(data));
		} else if (Action.SCATTER_CHART.equals(toolbarAction)) {
			doScatterChart(getSelection(data));
		} else if (Action.DOUGHNUT_CHART.equals(toolbarAction)) {
			doDoughnutChart(getSelection(data));
		} else if (Action.HYPERLINK.equals(toolbarAction)) {
			doHyperlink(getSelection(data));
		} else if (Action.INSERT_PICTURE.equals(toolbarAction)) {
			doBeforeInsertPicture(getSelection(data));
		} else if (Action.CLOSE_BOOK.equals(toolbarAction)) {
			doCloseBook();
		} else if (Action.FORMAT_CELL.equals(toolbarAction)) {
			doFormatCell(getSelection(data));
		} else if (Action.COLUMN_WIDTH.equals(toolbarAction)) {
			doColumnWidth(getSelection(data));
		} else if (Action.ROW_HEIGHT.equals(toolbarAction)) {
			doRowHeight(getSelection(data));
		} else if (Action.HIDE_COLUMN.equals(toolbarAction)) {
			doHideColumn(getSelection(data));
		} else if (Action.UNHIDE_COLUMN.equals(toolbarAction)) {
			doUnhideColumn(getSelection(data));
		} else if (Action.HIDE_ROW.equals(toolbarAction)) {
			doHideRow(getSelection(data));
		} else if (Action.UNHIDE_ROW.equals(toolbarAction)) {
			doUnhideRow(getSelection(data));
		} else if (Action.INSERT_FUNCTION.equals(toolbarAction)) {
			doInsertFunction(getSelection(data));
		}
	}
	
	/**
	 * @param selection
	 */
	public abstract void doInsertFunction(Rect selection);
	
	/**
	 * @param selection
	 */
	public void doHideRow(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Ranges
			.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight())
			.getRows()
			.setHidden(true);
		}
	}
	/**
	 * @param selection
	 */
	public void doUnhideRow(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Ranges
			.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight())
			.getRows()
			.setHidden(false);	
		}
	}
	/**
	 * @param selection
	 */
	public void doUnhideColumn(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Ranges
			.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight())
			.getColumns()
			.setHidden(false);	
		}
	}
	/**
	 * @param selection
	 */
	public void doHideColumn(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Ranges
			.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight())
			.getColumns()
			.setHidden(true);	
		}
	}
	/**
	 * @param selection
	 */
	public abstract void doColumnWidth(Rect selection);
	
	/**
	 * @param selection
	 */
	public abstract void doRowHeight(Rect selection);
	
	public void doCloseBook() {
		_spreadsheet.setSrc(null);
		clearClipboard();
		_insertPictureSelection = null;
		for (Action a : disabledActionOnBookClose) {
			_spreadsheet.setActionDisabled(true, a);
		}
	}
	
	public void doInsertPicture(UploadEvent evt) {
		if (_insertPictureSelection == null) {
			return;
		}
		
		final Media media = evt.getMedia();
		if (media instanceof AImage) {
			AImage image = (AImage)media;
			Ranges
			.range(_spreadsheet.getSelectedSheet())
			.addPicture(getClientAnchor(_insertPictureSelection.getTop(), _insertPictureSelection.getLeft(), 
					image.getWidth(), image.getHeight()), image.getByteData(), getImageFormat(image));
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
	
	/**
	 * @param selection
	 */
	public void doBeforeInsertPicture(Rect selection) {
		_insertPictureSelection = selection;
	}
	
	protected void setVerticalAlign(final short alignment, Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Utils.visitCells(sheet, selection, new CellVisitor(){
				@Override
				public void handle(CellVisitorContext context) {
					final short srcAlign = context.getVerticalAlignment();

					if (srcAlign != alignment) {
						CellStyle newStyle = context.cloneCellStyle();
						newStyle.setVerticalAlignment(alignment);
						context.getRange().setStyle(newStyle);
					}
				}});
		}
	}
	
	protected void setHorizontalAlign(final short alignment, Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Utils.visitCells(sheet, selection, new CellVisitor(){
				@Override
				public void handle(CellVisitorContext context) {
					final short srcAlign = context.getAlignment();

					if (srcAlign != alignment) {
						CellStyle newStyle = context.cloneCellStyle();
						newStyle.setAlignment(alignment);
						context.getRange().setStyle(newStyle);
					}
				}});
		}
	}
	
	protected Rect getSelection(Map data) {
		int tRow = (Integer) data.get("tRow");
		int bRow = (Integer) data.get("bRow");
		int lCol = (Integer) data.get("lCol");
		int rCol = (Integer) data.get("rCol");
		Rect r = new Rect(lCol, tRow, rCol, bRow);
		return r;
	}
	
	protected void syncClipboard() {
		if (_clipboard != null) {
			final Book srcBook = _clipboard.book;
			if (!srcBook.equals(_spreadsheet.getBook())) {
				clearClipboard();
			} else {
				final Worksheet srcSheet = _clipboard.sourceSheet;
				boolean validSheet = srcBook.getSheetIndex(srcSheet) >= 0;
				if (!validSheet) {
					_spreadsheet.setHighlight(null);
					clearClipboard();
				} else if (!srcSheet.equals(_spreadsheet.getSelectedSheet())) {
					_spreadsheet.setHighlight(null);
				} else {
					_spreadsheet.setHighlight(_clipboard.sourceRect);
				}
			}
		}
	}
	
	protected void syncAutoFilter() {
		final Worksheet worksheet = _spreadsheet.getSelectedSheet();
		boolean appliedFilter = false;
		AutoFilter af = worksheet.getAutoFilter();
		if (af != null) {
			final CellRangeAddress afrng = af.getRangeAddress();
			if (afrng != null) {
				int rowIdx = afrng.getFirstRow() + 1;
				for (int i = rowIdx; i <= afrng.getLastRow(); i++) {
					final Row row = worksheet.getRow(i);
					if (row != null && row.getZeroHeight()) {
						appliedFilter = true;
						break;
					}
				}	
			}
		}

		_spreadsheet.setActionDisabled(!appliedFilter, Action.CLEAR_FILTER);
		_spreadsheet.setActionDisabled(!appliedFilter, Action.REAPPLY_FILTER);
		
		if (!Objects.equals(_book, _spreadsheet.getBook())) {
			if (_book != null) {
				_book.unsubscribe(_bookListener);
			}
			_book = _spreadsheet.getBook();
			_book.subscribe(_bookListener);
		}
	}
	
	private void init() {
		_spreadsheet.addEventListener(Events.ON_SHEET_SELECT, new EventListener() {
			
			@Override
			public void onEvent(Event event) throws Exception {
				doSheetSelect();
			}
		});
		_spreadsheet.addEventListener(Events.ON_CTRL_KEY, new EventListener() {

			@Override
			public void onEvent(Event event) throws Exception {
				KeyEvent evt = (KeyEvent) event;
				doCtrlKey((KeyEvent) event);
			}
		});
		
		Upload upload = new Upload();
		upload.appendChild(_insertPicture = new Uploader());
		_insertPicture.setId(Action.INSERT_PICTURE.toString());
		
		_insertPicture.addEventListener(org.zkoss.zk.ui.event.Events.ON_UPLOAD, new EventListener() {
			
			@Override
			public void onEvent(Event event) throws Exception {
				doInsertPicture((UploadEvent)event);
			}
		});
		_spreadsheet.appendChild(upload);
	}
	
	/**
	 * Execute when user press key
	 * @param event
	 */
	public void doCtrlKey(KeyEvent event) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		Rect selection = event.getSelection();
		if (sheet == null || !validSelection(selection)) {
			return;
		}
		
		if (46 == event.getKeyCode()) {
			if (event.isCtrlKey())
				doClearStyle(selection);
			else
				doClearContent(selection);
			return;
		}
		if (false == event.isCtrlKey())
			return;
		
		char keyCode = (char) event.getKeyCode();
		switch (keyCode) {
		case 'X':
			doCut(selection);
			break;
		case 'C':
			doCopy(selection);
			break;
		case 'V':
			doPaste(selection);
			break;
		case 'D':
			doClearContent(selection);
			break;
		case 'B':
			doFontBold(selection);
			break;
		case 'I':
			doFontItalic(selection);
			break;
		case 'U':
			doFontUnderline(selection);
			break;
		}
	}
	
	/**
	 * Execute when user select sheet
	 */
	public void doSheetSelect() {
		syncClipboard();
		syncAutoFilter();
		
		//enable action button
		for (Action action : disabledActionOnBookClose) {
			_spreadsheet.setActionDisabled(false, action);
		}
		_spreadsheet.setActionDisabled(true, Action.SAVE_BOOK);
	}
	
	/**
	 * Bind the handler's target
	 * 
	 * @param spreadsheet
	 */
	public void bind(Spreadsheet spreadsheet) {
		if (_spreadsheet != spreadsheet) {
			_spreadsheet = spreadsheet;
			init();	
		}
	}
	
	/**
	 * Returns {@link Spreadsheet}
	 * @return
	 */
	public Spreadsheet getSpreadsheet() {
		return _spreadsheet;
	} 
	
	/**
	 * When user click Home pane
	 * 
	 * <p>
	 * Default: do nothing
	 */
	public void doHomePanel() {
	}
	
	/**
	 * Execute when user click formula panel
	 * 
	 * <p>
	 * Default: do nothing
	 */
	public void doFormulaPanel() {
	}
	
	/**
	 * Execute when user click insert panel
	 * 
	 * <p>
	 * Default: do nothing
	 */
	public void doInsertPanel() {
	}

	/**
	 * Execute when user click new book
	 */
	public abstract void doNewBook();
	
	/**
	 * Execute when user click save book
	 */
	public abstract void doSaveBook();
	
	/**
	 * Execute when user click export PDF
	 */
	public abstract void doExportPDF(Rect selection);
	
	public Clipboard getClipboard() {
		return _clipboard;
	}
	
	public void clearClipboard() {
		_clipboard = null;
	}
	
	/**
	 * Execute when user click copy
	 */
	public void doCopy(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			_clipboard = new Clipboard(Clipboard.Type.COPY, selection, sheet, _spreadsheet.getBook());
			_spreadsheet.setHighlight(selection);
		}
	}
	
	/**
	 * @param pasteType
	 * @param pasteOperation
	 * @param skipBlank
	 * @param transpose
	 */
	protected void doPasteImpl(Rect selection, int pasteType, int pasteOperation, boolean skipBlank, boolean transpose) {
		Book srcBook = _clipboard.book;
		Book targetBook = _spreadsheet.getBook();
		if (targetBook != null && targetBook.equals(srcBook)) {
			final Worksheet srcSheet = _clipboard.sourceSheet;
			boolean validSheet = srcBook.getSheetIndex(srcSheet) >= 0;
			if (!validSheet) {
				clearClipboard();
				return;
			}
			
			final Rect srcRect = _clipboard.sourceRect;
			Range rng = Utils.pasteSpecial(srcSheet,
					srcRect, 
					_spreadsheet.getSelectedSheet(), 
					selection.getTop(),
					selection.getLeft(),
					selection.getBottom(),
					selection.getRight(),
					pasteType, 
					pasteOperation, 
					skipBlank, transpose);
			
			if (_clipboard.type == Clipboard.Type.CUT) {
				Ranges
				.range(srcSheet, srcRect.getTop(), srcRect.getLeft(), srcRect.getBottom(), srcRect.getRight())
				.clearContents();
				
				clearStyleImp(srcRect, srcSheet);
				
				_clipboard = null;//clear used clipboard
				_spreadsheet.setHighlight(null);
			}
			
			if (rng != null) {
				_spreadsheet.setSelection(new Rect(rng.getColumn(), rng.getRow(), 
						rng.getLastColumn(), rng.getLastRow()));	
			}
		}
	}
	
	protected boolean validSelection(Rect selection) {
		return selection.getTop() >= 0 && selection.getLeft() >= 0
			&& selection.getBottom() >= 0 && selection.getRight() >= 0;
	}
	
	/**
	 * Execute when user click paste 
	 */
	public void doPaste(Rect selection) {
		if (_clipboard != null && validSelection(selection)) {
			doPasteImpl(selection, Range.PASTE_ALL, Range.PASTEOP_NONE, false, false);
		}
	}
	
	
	/**
	 * Execute when user click paste formula
	 */
	public void doPasteFormula(Rect selection) {
		if (_clipboard != null && validSelection(selection)) {
			doPasteImpl(selection, Range.PASTE_FORMULAS, Range.PASTEOP_NONE, false, false);
		}
	}
	
	/**
	 *  Execute when user click paste value
	 */
	public void doPasteValue(Rect selection) {
		if (_clipboard != null && validSelection(selection)) {
			doPasteImpl(selection, Range.PASTE_VALUES, Range.PASTEOP_NONE, false, false);
		}
	}
	
	/**
	 * Execute when user click paste all except border
	 */
	public void doPasteAllExceptBorder(Rect selection) {
		if (_clipboard != null && validSelection(selection)) {
			doPasteImpl(selection, Range.PASTE_ALL_EXCEPT_BORDERS, Range.PASTEOP_NONE, false, false);
		}
	}
	
	/**
	 * Execute when user click paste transpose
	 */
	public void doPasteTranspose(Rect selection) {
		if (_clipboard != null && validSelection(selection)) {
			doPasteImpl(selection, Range.PASTE_ALL, Range.PASTEOP_NONE, false, true);
		}
	}
	
	/**
	 * Execute when user click paste special
	 */
	public abstract void doPasteSpecial(Rect selection);
	
	/**
	 * Execute when user click cut
	 */
	public void doCut(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			_clipboard = new Clipboard(Clipboard.Type.CUT, selection, sheet, _spreadsheet.getBook());
			_spreadsheet.setHighlight(selection);
		}
	}
	
	/**
	 * Execute when user select font family
	 * 
	 * @param selection
	 */
	public void doFontFamily(String fontFamily, Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Utils.setFontFamily(sheet, selection, fontFamily);	
		}
	}
	
	/**
	 * Execute when user select font size
	 *  
	 * @param selection
	 */
	public void doFontSize(int fontSize, Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			short fontHeightInPoint = (short)(fontSize * 20);
			Utils.setFontHeight(sheet, selection, fontHeightInPoint);		
		}
	}
	
	private Font getCellFont(int row, int col) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		Cell cell = Utils.getOrCreateCell(sheet, row, col);
		return _spreadsheet.getBook().getFontAt(cell.getCellStyle().getFontIndex());
	}
	
	/**
	 * Execute when user click font bold
	 * 
	 * @param selection
	 */
	public void doFontBold(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			int row = selection.getTop();
			int col = selection.getLeft();
			boolean fontBold = Font.BOLDWEIGHT_BOLD ==  getCellFont(row, col).getBoldweight();
			Utils.setFontBold(sheet, selection, !fontBold);	
		}
	}
	
	/**
	 * Execute when user click font italic
	 * 
	 * @param selection
	 */
	public void doFontItalic(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			int row = selection.getTop();
			int col = selection.getLeft();
			Utils.setFontItalic(sheet, selection, !getCellFont(row, col).getItalic());	
		}
	}
	
	/**
	 * Execute when user click font strikethrough
	 * 
	 * @param selection
	 */
	public void doFontStrikeout(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			int row = selection.getTop();
			int col = selection.getLeft();
			Utils.setFontStrikeout(sheet, selection, !getCellFont(row, col).getStrikeout());	
		}
	}
	
	/**
	 * Execute when user click font underline
	 * 
	 * @param selection
	 */
	public void doFontUnderline(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			int row = selection.getTop();
			int col = selection.getLeft();
			boolean underline = Font.U_SINGLE == getCellFont(row, col).getUnderline();
			Utils.setFontUnderline(sheet, selection, underline ? Font.U_NONE : Font.U_SINGLE);	
		}
	}
	
	/**
	 * Returns the border color
	 * @return
	 */
	protected String getBorderColor() {
		return "#000000";
	}
	
	/**
	 * Execute when user click border
	 * 
	 * @param selection
	 */
	public void doBorder(String color, Rect selection) {
		doBorderBottom(color, selection);
	}
	
	/**
	 * Execute when user click bottom border
	 * 
	 * @param selection
	 */
	public void doBorderBottom(String color, Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Utils.setBorder(sheet, selection, 
					BookHelper.BORDER_EDGE_BOTTOM, BorderStyle.MEDIUM, color);
		}
	}
	
	/**
	 * Execute when user click top border
	 * 
	 * @param selection
	 */
	public void doBoderTop(String color, Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Utils.setBorder(sheet, selection, 
					BookHelper.BORDER_EDGE_TOP, BorderStyle.MEDIUM, color);	
		}
	}
	
	/**
	 * Execute when user click left border
	 * 
	 * @param selection
	 */
	public void doBorderLeft(String color, Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Utils.setBorder(sheet, selection, 
					BookHelper.BORDER_EDGE_LEFT, BorderStyle.MEDIUM, color);
		}
	}
	
	/**
	 * Execute when user click right border
	 * 
	 * @param selection
	 */
	public void doBorderRight(String color, Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Utils.setBorder(sheet, selection, 
					BookHelper.BORDER_EDGE_RIGHT, BorderStyle.MEDIUM, color);	
		}
	}
	
	/**
	 * Execute when user click no border
	 * 
	 * @param selection
	 */
	public void doBorderNo(String color, Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Utils.setBorder(sheet, selection, 
					BookHelper.BORDER_FULL, BorderStyle.NONE, color);	
		}
	}
	
	/**
	 * Execute when user click all border
	 * 
	 * @param selection
	 */
	public void doBorderAll(String color, Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Utils.setBorder(sheet, selection, 
					BookHelper.BORDER_FULL, BorderStyle.MEDIUM, color);	
		}
	}
	
	/**
	 * Execute when user click outside border
	 * 
	 * @param selection
	 */
	public void doBorderOutside(String color, Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Utils.setBorder(sheet, selection, 
					BookHelper.BORDER_OUTLINE, BorderStyle.MEDIUM, color);	
		}
	}
	
	/**
	 * Execute when user click inside border
	 * 
	 * @param selection
	 */
	public void doBorderInside(String color, Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Utils.setBorder(sheet, selection, 
					BookHelper.BORDER_INSIDE, BorderStyle.MEDIUM, color);	
		}
	}
	
	/**
	 * Execute when user click inside horizontal border
	 * 
	 * @param selection
	 */
	public void doBorderInsideHorizontal(String color, Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Utils.setBorder(sheet, selection, 
					BookHelper.BORDER_INSIDE_HORIZONTAL, BorderStyle.MEDIUM, color);	
		}
	}
	
	/**
	 * Execute when user click inside vertical border
	 * 
	 * @param selection
	 */
	public void doBorderInsideVertical(String color, Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Utils.setBorder(sheet, selection, 
					BookHelper.BORDER_INSIDE_VERTICAL, BorderStyle.MEDIUM, color);	
		}
	}
	
	/**
	 * Execute when user click font color 
	 * 
	 * @param color
	 * @param selection
	 */
	public void doFontColor(String color, Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Utils.setFontColor(sheet, selection, color);	
		}
	}
	
	/**
	 * Execute when user click fill color
	 * 
	 * @param color
	 * @param selection
	 */
	public void doFillColor(String color, Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Utils.setBackgroundColor(sheet, selection, color);	
		}
	}
	
	/**
	 * Execute when user click vertical align top
	 * 
	 * @param selection
	 */
	public void doVerticalAlignTop(Rect selection) {
		setVerticalAlign(CellStyle.VERTICAL_TOP, selection);
	}
	
	/**
	 * Execute when user click vertical align middle
	 * 
	 * @param selection
	 */
	public void doVerticalAlignMiddle(Rect selection) {
		setVerticalAlign(CellStyle.VERTICAL_CENTER, selection);
	}

	/**
	 * Execute when user click vertical align bottom
	 * 
	 * @param selection
	 */
	public void doVerticalAlignBottom(Rect selection) {
		setVerticalAlign(CellStyle.VERTICAL_BOTTOM, selection);
	}
	
	/**
	 * @param selection
	 */
	public void doHorizontalAlignLeft(Rect selection) {
		setHorizontalAlign(CellStyle.ALIGN_LEFT, selection);
	}
	
	/**
	 * @param selection
	 */
	public void doHorizontalAlignCenter(Rect selection) {
		setHorizontalAlign(CellStyle.ALIGN_CENTER, selection);
	}
	
	/**
	 * @param selection
	 */
	public void doHorizontalAlignRight(Rect selection) {
		setHorizontalAlign(CellStyle.ALIGN_RIGHT, selection);
	}
	
	/**
	 * @param selection
	 */
	public void doWrapText(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			int row = selection.getTop();
			int col = selection.getLeft();
			final boolean wrapText = !Utils.getOrCreateCell(sheet, row, col).getCellStyle().getWrapText();
			Utils.setWrapText(sheet, selection, wrapText);	
		}
	}
	
	/**
	 * @param selection
	 */
	public void doMergeAndCenter(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			int tRow = selection.getTop();
			int lCol = selection.getLeft();
			int bRow = selection.getBottom();
			int rCol = selection.getRight();
			
			boolean merged = false;
			MergeMatrixHelper mergeHelper = _spreadsheet.getMergeMatrixHelper(sheet);
			for (int r = tRow; r <= bRow; r++) {
				for (int c = lCol; c <= rCol; c++) {
					MergedRect rect = mergeHelper.getMergeRange(r, c);
					if (rect != null) {
						merged = true;
						break;
					}
				}
			}
			
			Range range = Ranges.range(sheet, tRow, lCol, bRow, rCol);
			if (merged) {
				range.unMerge();
			} else {
				range.merge(false);
				doHorizontalAlignCenter(selection);	
			}
		}
	}
	
	/**
	 * @param selection
	 */
	public void doMergeAcross(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Ranges
			.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight())
			.merge(true);	
		}
	}
	
	/**
	 * @param selection
	 */
	public void doMergeCell(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Ranges
			.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight())
			.merge(false);	
		}
	}
	
	/**
	 * @param selection
	 */
	public void doUnmergeCell(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Ranges
			.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight())
			.unMerge();	
		}
	}
	
	/**
	 * @param selection
	 */
	public void doShiftCellRight(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Ranges
			.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight())
			.insert(Range.SHIFT_RIGHT, Range.FORMAT_RIGHTBELOW);	
		}
	}
	
	/**
	 * @param selection
	 */
	public void doShiftCellDown(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Ranges
			.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight())
			.insert(Range.SHIFT_DOWN, Range.FORMAT_LEFTABOVE);	
		}
	}
	
	/**
	 * @param selection
	 */
	public void doInsertSheetRow(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Ranges
			.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight())
			.getRows()
			.insert(Range.SHIFT_DOWN, Range.FORMAT_LEFTABOVE);	
		}
	}
	
	/**
	 * @param selection
	 */
	public void doInsertSheetColumn(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Ranges
			.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight())
			.getColumns()
			.insert(Range.SHIFT_RIGHT, Range.FORMAT_RIGHTBELOW);	
		}
	}
	
	/**
	 * @param selection
	 */
	public void doShiftCellLeft(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Ranges
			.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight())
			.delete(Range.SHIFT_LEFT);	
		}
	}
	
	/**
	 * @param selection
	 */
	public void doShiftCellUp(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Ranges
			.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight())
			.delete(Range.SHIFT_UP);	
		}
	}
	
	/**
	 * @param selection
	 */
	public void doDeleteSheetRow(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Ranges
			.range(sheet, selection.getTop(), 0, selection.getBottom(), 0)
			.getRows()
			.delete(Range.SHIFT_UP);	
		}
	}
	
	/**
	 * @param selection
	 */
	public void doDeleteSheetColumn(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Ranges
			.range(sheet, 0, selection.getLeft(), 0, selection.getRight())
			.getColumns()
			.delete(Range.SHIFT_LEFT);	
		}
	}

	protected void clearStyleImp(Rect selection, Worksheet worksheet) {
		final CellStyle defaultStyle = worksheet.getBook().createCellStyle();
		Ranges
		.range(worksheet, selection.getTop(), selection.getLeft(),selection.getBottom(), selection.getRight())
		.setStyle(defaultStyle);
	}
	
	/**
	 * Execute when user click clear style
	 */
	public void doClearStyle(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			clearStyleImp(selection, sheet);	
		}
	}
	
	/**
	 * Execute when user click clear content
	 */
	public void doClearContent(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Ranges
			.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight())
			.clearContents();	
		}
	}
	
	/**
	 * Execute when user click clear all
	 */
	public void doClearAll(Rect selection) {
		doClearContent(selection);
		doClearStyle(selection);
	}
	
	/**
	 * @param selection
	 */
	public void doSortAscending(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Utils.sort(sheet, selection,
					null, null, null, false, false, false);	
		}
	}
	
	/**
	 * @param selection
	 */
	public void doSortDescending(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Utils.sort(sheet, selection,
					null, new boolean[] { true }, null, false, false, false);	
		}
	}
	
	/**
	 * @param selection
	 */
	public abstract void doCustomSort(Rect selection);
	
	/**
	 * @param selection
	 */
	public void doFilter(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			Ranges
			.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight())
			.autoFilter();	
		}
	}
	
	/**
	 * @param selection
	 */
	public void doClearFilter() {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null) {
			Ranges.range(sheet).showAllData();	
		}
	}
	
	/**
	 * @param selection
	 */
	public void doReapplyFilter() {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null) {
			Ranges.range(sheet).applyFilter();	
		}
	}
	
	/**
	 * 
	 */
	public void doProtectSheet() {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null) {
			Ranges.range(sheet).protectSheet(sheet.getProtect() ? null : "");	
		}
	}
	
	/**
	 * 
	 */
	public void doGridlines() {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null) {
			Ranges.range(sheet).setDisplayGridlines(!sheet.isDisplayGridlines());	
		}
	}
	
	/**
	 * @param selection
	 */
	public void doColumnChart(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			int row = selection.getTop();
			int col = selection.getLeft();
			ChartData data = fillCategoryData(new XSSFColumnChartData(), selection);
			
			Ranges.range(sheet)
			.addChart(getClientAnchor(row, col, 600, 300), data, ChartType.Column, ChartGrouping.STANDARD, LegendPosition.RIGHT);	
		}
	}
	
	/**
	 * @param selection
	 */
	public void doColumnChart3D(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			int row = selection.getTop();
			int col = selection.getLeft();
			ChartData data = fillCategoryData(new XSSFColumn3DChartData(), selection);
			
			Ranges.range(sheet)
			.addChart(getClientAnchor(row, col, 600, 300), data, ChartType.Column3D, ChartGrouping.STANDARD, LegendPosition.RIGHT);	
		}
	}
	/**
	 * @param selection
	 */
	public void doLineChart(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			int row = selection.getTop();
			int col = selection.getLeft();
			ChartData data = fillCategoryData(new XSSFLineChartData(), selection);
			
			Ranges.range(sheet)
			.addChart(getClientAnchor(row, col, 600, 300), data, ChartType.Line, ChartGrouping.STANDARD, LegendPosition.RIGHT);	
		}
	}
	/**
	 * @param selection
	 */
	public void doLineChart3D(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			int row = selection.getTop();
			int col = selection.getLeft();
			ChartData data = fillCategoryData(new XSSFLine3DChartData(), selection);
			
			Ranges.range(sheet)
			.addChart(getClientAnchor(row, col, 600, 300), data, ChartType.Line3D, ChartGrouping.STANDARD, LegendPosition.RIGHT);	
		}
	}
	/**
	 * @param selection
	 */
	public void doPieChart(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			int row = selection.getTop();
			int col = selection.getLeft();
			ChartData data = fillCategoryData(new XSSFPieChartData(), selection);
			
			Ranges.range(sheet)
			.addChart(getClientAnchor(row, col, 600, 300), data, ChartType.Pie, ChartGrouping.STANDARD, LegendPosition.RIGHT);	
		}
	}
	/**
	 * @param selection
	 */
	public void doPieChart3D(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			int row = selection.getTop();
			int col = selection.getLeft();
			ChartData data = fillCategoryData(new XSSFPie3DChartData(), selection);
			
			Ranges.range(sheet)
			.addChart(getClientAnchor(row, col, 600, 300), data, ChartType.Pie3D, ChartGrouping.STANDARD, LegendPosition.RIGHT);	
		}
	}
	/**
	 * @param selection
	 */
	public void doBarChart(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			int row = selection.getTop();
			int col = selection.getLeft();
			ChartData data = fillCategoryData(new XSSFBarChartData(), selection);
			
			Ranges.range(sheet)
			.addChart(getClientAnchor(row, col, 600, 300), data, ChartType.Bar, ChartGrouping.STANDARD, LegendPosition.RIGHT);	
		}
	}
	/**
	 * @param selection
	 */
	public void doBarChart3D(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			int row = selection.getTop();
			int col = selection.getLeft();
			ChartData data = fillCategoryData(new XSSFBar3DChartData(), selection);
			
			Ranges.range(sheet)
			.addChart(getClientAnchor(row, col, 600, 300), data, ChartType.Bar3D, ChartGrouping.STANDARD, LegendPosition.RIGHT);	
		}
	}
	/**
	 * @param selection
	 */
	public void doAreaChart(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			int row = selection.getTop();
			int col = selection.getLeft();
			ChartData data = fillCategoryData(new XSSFAreaChartData(), selection);
			
			Ranges.range(sheet)
			.addChart(getClientAnchor(row, col, 600, 300), data, ChartType.Area, ChartGrouping.STANDARD, LegendPosition.RIGHT);	
		}
	}
	/**
	 * @param selection
	 */
	public void doScatterChart(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			int row = selection.getTop();
			int col = selection.getLeft();
			ChartData data = fillXYData(new XSSFScatChartData(), selection);
			
			Ranges.range(sheet)
			.addChart(getClientAnchor(row, col, 600, 300), data, ChartType.Scatter, ChartGrouping.STANDARD, LegendPosition.RIGHT);	
		}
	}
	/**
	 * @param selection
	 */
	public void doDoughnutChart(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if (sheet != null && validSelection(selection)) {
			int row = selection.getTop();
			int col = selection.getLeft();
			ChartData data = fillCategoryData(new XSSFDoughnutChartData(), selection);
			
			Ranges.range(sheet)
			.addChart(getClientAnchor(row, col, 600, 300), data, ChartType.Doughnut, ChartGrouping.STANDARD, LegendPosition.RIGHT);	
		}
	}
	
	private boolean isQualifiedCell(Cell cell) {
		if (cell == null)
			return true;
		int cellType = cell.getCellType();
		return cellType == Cell.CELL_TYPE_NUMERIC || 
			cellType == Cell.CELL_TYPE_FORMULA || 
			cellType == Cell.CELL_TYPE_BLANK;
	}
	
	private Rect getChartDataRange(Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
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
	
	protected ChartData fillXYData(XYData data, Rect selection) {
		final Worksheet sheet = _spreadsheet.getSelectedSheet();
		
		Rect rect = getChartDataRange(selection);
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
				String startCell = _spreadsheet.getColumntitle(lCol) + _spreadsheet.getRowtitle(rowIdx);
				String endCell = _spreadsheet.getColumntitle(rCol) + _spreadsheet.getRowtitle(selection.getBottom());
				horValues = DataSources.fromNumericCellRange(sheet, CellRangeAddress.valueOf(startCell + ":" + endCell));
			}
			//find values
			int i = 1;
			for (int c = colIdx; c <= selection.getRight(); c++) {
				//find title
				String title = null;
				int row = rowIdx - 1;
				if (row >= selection.getTop()) {
					title = "" + Ranges.range(sheet, selection.getTop(), c, row, c).getText().toString();
				}
				titles.add(title == null ? null : DataSources.fromString(title));
				
				String startCell = _spreadsheet.getColumntitle(c) + _spreadsheet.getRowtitle(rowIdx);
				String endCell = _spreadsheet.getColumntitle(c) + _spreadsheet.getRowtitle(selection.getBottom());
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
				String startCell = _spreadsheet.getColumntitle(colIdx) + _spreadsheet.getRowtitle(tRow);
				String endCell = _spreadsheet.getColumntitle(selection.getRight()) + _spreadsheet.getRowtitle(tRow);
				horValues = DataSources.fromNumericCellRange(sheet, CellRangeAddress.valueOf(startCell + ":" + endCell));
			}
			//find values
			int i = 1;
			for (int r = rowIdx; r <= selection.getBottom(); r++) {
				//find title
				String title = null;
				int col = colIdx - 1;
				if (col >= selection.getLeft()) {
					title = "" + Ranges.range(sheet, r, selection.getLeft(), r, col).getText().toString();
				}
				titles.add(title == null ? null : DataSources.fromString(title));
				
				String startCell = _spreadsheet.getColumntitle(colIdx) + _spreadsheet.getRowtitle(r);
				String endCell = _spreadsheet.getColumntitle(selection.getRight()) + _spreadsheet.getRowtitle(r);
				values.add(DataSources.fromNumericCellRange(sheet, CellRangeAddress.valueOf(startCell + ":" + endCell)));
			}
		}
		
		for (int i = 0; i < values.size(); i++) {
			data.addSerie(titles.get(i), horValues, values.get(i));
		}
		return data;
	}
	
	protected CategoryData fillCategoryData(CategoryData data, Rect selection) {
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		Rect rect = getChartDataRange(selection);
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
				String startCell = _spreadsheet.getColumntitle(selection.getLeft()) + _spreadsheet.getRowtitle(rowIdx);
				String endCell = _spreadsheet.getColumntitle(col) + _spreadsheet.getRowtitle(selection.getBottom());
				cats = DataSources.fromStringCellRange(sheet, CellRangeAddress.valueOf(startCell + ":" + endCell));
			}
			//find value, by column
			int i = 1;
			for (int c = colIdx; c <= selection.getRight(); c++) {
				//find title
				String title = null;
				int row = rowIdx - 1;
				if (row >= selection.getTop()) {
					title = "" + Ranges.range(sheet, selection.getTop(), c, row, c).getText().toString();
				}
				titles.add(title == null ? null : DataSources.fromString(title));
				
				String startCell = _spreadsheet.getColumntitle(c) + _spreadsheet.getRowtitle(rowIdx);
				String endCell = _spreadsheet.getColumntitle(c) + _spreadsheet.getRowtitle(selection.getBottom());
				ChartDataSource<Number> val = DataSources.fromNumericCellRange(sheet, CellRangeAddress.valueOf(startCell + ":" + endCell));
				vals.add(val);
			}
		} else { //catalog by column, value by row
			//find catalog
			int row = rowIdx - 1;
			if (row >= selection.getTop()) {
				String startCell = _spreadsheet.getColumntitle(colIdx) + _spreadsheet.getRowtitle(row);
				String endCell = _spreadsheet.getColumntitle(selection.getRight()) + _spreadsheet.getRowtitle(row);
				cats = DataSources.fromStringCellRange(sheet, CellRangeAddress.valueOf(startCell + ":" + endCell));
			}
			
			//find value
			int i = 1;
			for (int r = rowIdx; r <= selection.getBottom(); r++) {
				//find title
				String title = null;
				int col = colIdx - 1;
				if (col >= selection.getLeft()) {
					title = "" + Ranges.range(sheet, r, selection.getLeft(), r, col).getText().toString();
				}
				titles.add(title == null ? null : DataSources.fromString(title));
				
				String startCell = _spreadsheet.getColumntitle(colIdx) + _spreadsheet.getRowtitle(r);
				String endCell = _spreadsheet.getColumntitle(selection.getRight()) + _spreadsheet.getRowtitle(r);
				ChartDataSource<Number> val = DataSources.fromNumericCellRange(sheet, CellRangeAddress.valueOf(startCell + ":" + endCell));
				vals.add(val);
			}
		}
		
		for (int i = 0; i < vals.size(); i++) {
			data.addSerie(titles.get(i), cats, vals.get(i));
		}
		return data;
	}
	
	protected ClientAnchor getClientAnchor(int row, int col, int widgetWidth, int widgetHeight) {
		HeaderPositionHelper rowSizeHelper = (HeaderPositionHelper) _spreadsheet.getAttribute("_rowCellSize");
		HeaderPositionHelper colSizeHelper = (HeaderPositionHelper) _spreadsheet.getAttribute("_colCellSize");
		
		int lCol = col;
		int tRow = row;
		int rCol = lCol;
		int bRow = tRow;
		int offsetWidth = 0;
		int offsetHeight = 0;
		for (int r = tRow; r < _spreadsheet.getMaxrows(); r++) {
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
		for (int c = lCol; c < _spreadsheet.getMaxcolumns(); c++) {
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
	
	/** convert pixel to EMU */
	private static int pxToEmu(int px) {
		return (int) Math.round(((double)px) * 72 * 20 * 635 / 96); //assume 96dpi
	}
	
	/**
	 * 
	 */
	public abstract void doHyperlink(Rect selection);
	
	/**
	 * @param selection
	 */
	public abstract void doFormatCell(Rect selection);
	
	private static <T> T checkNotNull(String message, T t) {
		if (t == null) {
			throw new NullPointerException(message);
		}
		return t;
	}
	
	/**
	 * Used for copy & paste function
	 * 
	 * @author sam
	 */
	public static class Clipboard {
		public enum Type {
			COPY,
			CUT
		}
		
		public final Type type;
		public final Book book;
		public final Rect sourceRect;
		public final Worksheet sourceSheet;
		
		public Clipboard(Type type, Rect sourceRect, Worksheet sourceSheet, Book book) {
			this.type = checkNotNull("Clipboard's type cannot be null", type);
			this.book = checkNotNull("Clipboard's book cannot be null", book);
			this.sourceRect = checkNotNull("Clipboard's sourceRect cannot be null", sourceRect);
			this.sourceSheet = checkNotNull("Clipboard's sourceSheet cannot be null", sourceSheet);
		}
	}
}
