/* MainWindow.java

{{IS_NOTE
	Purpose:

	Description:

	History:
		Dec 20, 2007 12:40:46 PM     2007, Created by Dennis.Chenf
		Jan 1, 2009 modified by kinda lu
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

 */
package org.zkoss.zss.app;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.ArrayDeque;
import java.util.Iterator;
import java.util.List;

import org.omg.Messaging.SyncScopeHelper;
import org.zkoss.lang.Library;
import org.zkoss.poi.ss.usermodel.AutoFilter;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.FilterColumn;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.poi.ss.util.CellReference;
import org.zkoss.util.resource.Labels;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Path;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.event.KeyEvent;
import org.zkoss.zk.ui.event.SizeEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.cell.CellHelper;
import org.zkoss.zss.app.cell.EditHelper;
import org.zkoss.zss.app.ctrl.HeaderSizeCtrl;
import org.zkoss.zss.app.ctrl.RenameSheetCtrl;
import org.zkoss.zss.app.file.FileHelper;
import org.zkoss.zss.app.formula.FormulaMetaInfo;
import org.zkoss.zss.app.sheet.SheetHelper;
import org.zkoss.zss.app.sort.SortSelector;
import org.zkoss.zss.app.zul.CellContext;
import org.zkoss.zss.app.zul.CellMenupopup;
import org.zkoss.zss.app.zul.ColumnHeaderMenupopup;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.app.zul.EditMenu;
import org.zkoss.zss.app.zul.FileMenu;
import org.zkoss.zss.app.zul.FormatMenu;
import org.zkoss.zss.app.zul.FormulaEditor;
import org.zkoss.zss.app.zul.InsertMenu;
import org.zkoss.zss.app.zul.RowHeaderMenupopup;
import org.zkoss.zss.app.zul.Sheets;
import org.zkoss.zss.app.zul.ViewMenu;
import org.zkoss.zss.app.zul.Zssapp;
import org.zkoss.zss.app.zul.Zssapps;
import org.zkoss.zss.app.zul.ctrl.CellStyleContextEvent;
import org.zkoss.zss.app.zul.ctrl.CellStyleCtrlPanel;
import org.zkoss.zss.app.zul.ctrl.DesktopCellStyleContext;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zss.app.zul.ctrl.SSRectCellStyle;
import org.zkoss.zss.app.zul.ctrl.SSWorkbookCtrl;
import org.zkoss.zss.app.zul.ctrl.StyleModification;
import org.zkoss.zss.app.zul.ctrl.WorkbenchCtrl;
import org.zkoss.zss.app.zul.ctrl.WorkbookCtrl;
import org.zkoss.zss.engine.event.SSDataEvent;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.ui.Position;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.CellEvent;
import org.zkoss.zss.ui.event.CellMouseEvent;
import org.zkoss.zss.ui.event.CellSelectionEvent;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.event.HeaderEvent;
import org.zkoss.zss.ui.event.HeaderMouseEvent;
import org.zkoss.zss.ui.impl.MergeMatrixHelper;
import org.zkoss.zss.ui.impl.MergedRect;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Borderlayout;
import org.zkoss.zul.Checkbox;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.Comboitem;
import org.zkoss.zul.Div;
import org.zkoss.zul.Menu;
import org.zkoss.zul.Menuitem;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.North;
import org.zkoss.zul.Toolbarbutton;
import org.zkoss.zul.Window;

/**
 * @author Peter Kuo
 * @modify kinda lu
 */
public class MainWindowCtrl extends GenericForwardComposer implements WorkbenchCtrl {

	private static final long serialVersionUID = 1;
	
	private final static String KEY_PDF = "org.zkoss.zss.app.exportToPdf";
	private final static String KEY_HTML = "org.zkoss.zss.app.exportToHtml";
	
	static int event_x = 200;
	static int event_y = 200;

	int lastRow = 0;
	int lastCol = 0;
	boolean isFreezeRow = false;
	boolean isFreezeColumn = false;
	
	Spreadsheet spreadsheet;
	Sheets sheets;
	
	// For fast Icon
	boolean isMergeCell = false;

	int chartKey = 0;

	RangeHelper rangeh;

	Window mainWin;
	/*Menus*/
	FileMenu fileMenu;
	EditMenu editMenu;
	ViewMenu viewMenu;
	FormatMenu formatMenu;
	InsertMenu insertMenu;
	ColumnHeaderMenupopup columnHeaderMenupopup;
	RowHeaderMenupopup rowHeaderMenupopup;
	/*dropdown button menu*/
	Menuitem filter;
	Menuitem clearFilter;
	Menuitem reapplyFilter;
	
	CellContext cellContext;
	Menu insertImageMenu;
	Menu insertPieChart;
	
	/* Toolbar buttons */
	Borderlayout topToolbars;
	Div toolbarMask;
	Dropdownbutton pasteDropdownBtn;
	Dropdownbutton sortDropdownBtn;
	Toolbarbutton saveBtn;
	Toolbarbutton closeBtn;
	CellStyleCtrlPanel fontCtrlPanel;
	Toolbarbutton exportToPDFBtn;
	Toolbarbutton insertHyperlinkBtn;
	Toolbarbutton mergeCellBtn;
	Toolbarbutton insertChartBtn;
	Toolbarbutton cutBtn;
	Toolbarbutton copyBtn;
	Checkbox gridlinesCheckbox;
	Checkbox protectSheet;
	CellMenupopup cellMenupopup;
	
	/* Formula bar */
	North formulaBar;
	Combobox focusPosition;
	Comboitem lbpos;
	Toolbarbutton insertFormulaBtn;
	FormulaEditor formulaEditor;
	Div extendEditorBar;
	Toolbarbutton extendEditorBtn;
	boolean _formulaBarOpen = false; //whether extends height of the formulaBar
	private final static String FORMULA_BAR_ARROW_DOWN_IMAGE_SRC = "~./zssapp/image/arrow-270-small.png";
	private final static String FORMULA_BAR_ARROW_UP_IMAGE_SRC = "~./zssapp/image/arrow-090-small.png";
	
	/*Dialog*/
	Dialog _insertFormulaDialog;
	Dialog _insertHyperlinkDialog;
	Dialog _pasteSpecialDialog;
	Dialog _composeFormulaDialog;
	Dialog _formatNumberDialog;
	Dialog _saveFileDialog;
	Dialog _headerSizeDialog;
	Dialog _renameSheetDialog;
	Dialog _openFileDialog;
	Dialog _importFileDialog;
	Dialog _customSortDialog;
	Dialog _exportToPdfDialog;
	Dialog _exportToHtmlDialog;
	Dialog _autoFilterDialog;

	public Spreadsheet getSpreadsheet() {
		return spreadsheet;
	}

	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		//TODO: replace this mechanism
		initZssappComponents();
		init();
		rangeh = new RangeHelper(spreadsheet);
	}
	
	public void onCreate() {
		getDesktopWorkbenchContext().fireWorkbookChanged();
	}
	
	//TODO: remove this mechanism
	private void initZssappComponents() {
		Zssapps.bindSpreadsheet(spreadsheet, this);
	}
	
	public void init() {
		//Note. setSrcName will set spreadsheet's src name, but not the book name
		// if setSrc will init a book, then setSrcName only change the src name, 
		// if setSrc again with the first same book, the book will register two listener 
		//spreadsheet.setSrcName("Untitled");
		final DesktopWorkbenchContext workbenchContext = getDesktopWorkbenchContext();
		workbenchContext.doTargetChange(new SSWorkbookCtrl(spreadsheet));
		workbenchContext.setWorkbenchCtrl(this);
		workbenchContext.addEventListener(Consts.ON_SHEET_REFRESH, new EventListener(){
			public void onEvent(Event event) throws Exception {
				sheets.redraw();
			}
		});
		if (!FileHelper.hasSavePermission())
			saveBtn.setVisible(false);
		workbenchContext.addEventListener(Consts.ON_WORKBOOK_SAVED,	new EventListener() {
			public void onEvent(Event event) throws Exception {
				if (!FileHelper.hasSavePermission())
					return;
				setSaveButtonState(true);
			}
		});

		workbenchContext.addEventListener(Consts.ON_WORKBOOK_CHANGED, new EventListener() {
			public void onEvent(Event event) throws Exception {
				formulaEditor.setValue(null);
				boolean isOpen = spreadsheet.getBook() != null;
				formulaEditor.setDisabled(!isOpen);
				focusPosition.setDisabled(!isOpen);
				insertFormulaBtn.setDisabled(!isOpen);
				toolbarMask.setVisible(!isOpen);
				closeBtn.setVisible(isOpen);
				setSaveButtonState(isOpen ? true : null);

				gridlinesCheckbox.setChecked(isOpen && spreadsheet.getSelectedSheet().isDisplayGridlines());
				protectSheet.setChecked(isOpen && spreadsheet.getSelectedSheet().getProtect());
				protectSheet.setDisabled(!isOpen);
				if (isOpen) {
					sheets.redraw();
				} else if (!isOpen) {
					sheets.clear();
				}
				
				//TODO: provide clip board interface, to allow save cut, copy, high light info
				//use set setHighlight null can cancel selection, but need to re-store selection when select same sheet again
				spreadsheet.setHighlight(null);
				
				if (isOpen) {
					getCellStyleContext().doTargetChange(
							new SSRectCellStyle(Utils.getOrCreateCell(spreadsheet.getSelectedSheet(), 0, 0), spreadsheet));
				}
			}
		});
		workbenchContext.addEventListener(Consts.ON_SHEET_CHANGED, new EventListener() {
			public void onEvent(Event event) throws Exception {
				gridlinesCheckbox.setChecked(spreadsheet.getSelectedSheet().isDisplayGridlines());
				protectSheet.setChecked(spreadsheet.getSelectedSheet().getProtect());
				syncAutoFilterStatus();
			}
		});
		workbenchContext.addEventListener(Consts.ON_SHEET_CONTENTS_CHANGED,  new EventListener(){
			public void onEvent(Event event) throws Exception {
				doContentChanged();
			}}
		);
		//TODO: remove to WorkbookCtrl
		workbenchContext.addEventListener(Consts.ON_SHEET_MERGE_CELL, new EventListener() {
			public void onEvent(Event event) throws Exception {
				onMergeCellClick(null);
			}
		});
		workbenchContext.addEventListener(Consts.ON_SHEET_INSERT_FORMULA, new EventListener() {
			public void onEvent(Event event) throws Exception {
				String formula = (String)event.getData();
				Rect rect = spreadsheet.getSelection();
				Range rng = Ranges.range(spreadsheet.getSelectedSheet(), rect.getTop(), rect.getLeft());
				rng.setEditText(formula);
				formulaEditor.setText(rng.getEditText());
			}
		});
		workbenchContext.getWorkbookCtrl().addBookEventListener(new EventListener() {
			public void onEvent(Event event) throws Exception {
				String evtName = event.getName();
				if (evtName == SSDataEvent.ON_CONTENTS_CHANGE) {
					doContentChanged();
				} else if (evtName == SSDataEvent.ON_PROTECT_SHEET) {
					protectSheet.setChecked(spreadsheet.getSelectedSheet().getProtect());
				} else if (evtName == SSDataEvent.ON_BTN_CHANGE) {
					syncAutoFilterStatus();
				}
			}
		});

		// ADD Event Listener
		spreadsheet.addEventListener(Events.ON_CELL_FOUCSED,
				new EventListener() {
					public void onEvent(Event event) throws Exception {
						doFocusedEvent((CellEvent) event);
					}
				});

		spreadsheet.addEventListener(Events.ON_CELL_RIGHT_CLICK,
				new EventListener() {
					public void onEvent(Event event) throws Exception {
						doMouseEvent((CellMouseEvent) event);
					}
				});

		spreadsheet.addEventListener(Events.ON_HEADER_RIGHT_CLICK,
				new EventListener() {
					public void onEvent(Event event) throws Exception {
						doHeaderMouseEvent((HeaderMouseEvent) event);
					}
				});

		spreadsheet.addEventListener(Events.ON_CELL_SELECTION,
				new EventListener() {
					public void onEvent(Event event) throws Exception {
						doSelectionEvent((CellSelectionEvent) event);
					}
				});
		spreadsheet.addEventListener(Events.ON_CELL_RIGHT_CLICK, 
				new EventListener() {
					public void onEvent(Event event) throws Exception {
						CellMouseEvent evt = (CellMouseEvent)event;
						int clientX = evt.getClientx();
						int clientY = evt.getClienty();
						cellContext.setLeft(Integer.toString(clientX + 5) + "px");
						cellContext.setTop(Integer.toString(clientY - 100) + "px");
						cellContext.doPopup();
					}
				});
		spreadsheet.addEventListener(Events.ON_START_EDITING, 
				new EventListener() {
					public void onEvent(Event event) throws Exception {
						EditHelper.clearCutOrCopy(spreadsheet);
					}
				});
		syncAutoFilterStatus();
		
	}
	private void setSaveButtonState(Boolean saved) {
		if (saved == null) {
			saveBtn.setVisible(false);
			return;
		}
		
		saveBtn.setVisible(true);
		if (FileHelper.hasSavePermission()) {
			saveBtn.setDisabled(saved);
			saveBtn.setTooltiptext(saved ? Labels.getLabel("file.saved") : Labels.getLabel("file.saveNow"));
		}
	}

	public void doContentChanged() {
		setSaveButtonState(false);
		Worksheet seldSheet = spreadsheet.getSelectedSheet();
		Rect seld =  spreadsheet.getSelection();
		int row = seld.getTop();
		int col = seld.getLeft();
		Cell cell = Utils.getCell(seldSheet, row, col);
		if (cell != null) {
			getCellStyleContext().doTargetChange(new SSRectCellStyle(cell, spreadsheet));
			formulaEditor.setText(Ranges.range(seldSheet, row, col).getEditText());
		}
	}
	public void onClick$saveBtn() {
		DesktopWorkbenchContext workbench = getDesktopWorkbenchContext();
		if (workbench.getWorkbookCtrl().hasFileExtentionName()) {
			workbench.getWorkbookCtrl().save();
			workbench.fireWorkbookSaved();
		} else
			workbench.getWorkbenchCtrl().openSaveFileDialog();
	}
	public void onClick$exportToPDFBtn() {
		openExportPdfDialog();
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}
	public void onClick$pasteDropdownBtn() {
		getDesktopWorkbenchContext().getWorkbookCtrl().pasteSelection();
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}
	public void onDropdown$pasteDropdownBtn() {
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}
	public void onDropdown$sortDropdownBtn() {
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}
	public void onPasteSelector(ForwardEvent event) {
		String val = (String) event.getData();
		if ("pasteSpecial".equals(val))
			openPasteSpecialDialog();
		else
			EditHelper.onPasteEventHandler(spreadsheet, val);
			
	}
	public void onClick$insertFormulaBtn() {
		openInsertFormulaDialog();
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}
	private int _formulaBarSize = 26 + LINE_HEIGHT;
	private int toIntsize(String sz) {
		if (sz.endsWith("px")) {
			return Integer.valueOf(sz.substring(0, sz.length() - 2)).intValue();
		}
		return 26; //default
	}
	public void  onClick$extendEditorBtn() {
		_formulaBarOpen = !_formulaBarOpen;
		if (!_formulaBarOpen) {
			_formulaBarSize = toIntsize(formulaBar.getSize());
		}
		extendEditorBtn.setImage(
				_formulaBarOpen ? FORMULA_BAR_ARROW_UP_IMAGE_SRC : FORMULA_BAR_ARROW_DOWN_IMAGE_SRC);
		adjustFormulaBarHeight();
	}
	private static final int LINE_HEIGHT = 19;
	private void adjustFormulaBarHeight() {
		final int extra = _formulaBarSize - formulaBar.getMinsize();
		if (extra <= 0) {
			_formulaBarSize = formulaBar.getMinsize() + LINE_HEIGHT;
		} else {
			final int rest = extra % LINE_HEIGHT;
			_formulaBarSize = formulaBar.getMinsize() + (rest > 0 ? extra + (LINE_HEIGHT - rest) : extra);
		}
		String size = (_formulaBarOpen ? _formulaBarSize : formulaBar.getMinsize()) + "px";
		formulaBar.setSize(size);
		extendEditorBar.setHeight(size);
	}
	public void onSize$formulaBar(SizeEvent evt) {
		int height = toIntsize(evt.getHeight());
		if (height <= formulaBar.getMinsize()) {
			_formulaBarOpen = false;
		} else {
			_formulaBarSize = height;
			_formulaBarOpen = true;
		}
		extendEditorBtn.setImage(
				_formulaBarOpen ? FORMULA_BAR_ARROW_UP_IMAGE_SRC : FORMULA_BAR_ARROW_DOWN_IMAGE_SRC);
		adjustFormulaBarHeight();
	}
	public void onSortSelector(ForwardEvent event) {
		//TODO: replace forward event
		String param = (String) event.getData();
		if (param == null)
			return;
		if (param.equals(Labels.getLabel("sort.custom"))) {
			openCustomSortDialog();
		} else {
			if (SortSelector.getSortOrder(param))
				CellHelper.sortAscending(spreadsheet.getSelectedSheet(), SheetHelper.getSpreadsheetMaxSelection(spreadsheet));
			else
				CellHelper.sortDescending(spreadsheet.getSelectedSheet(), SheetHelper.getSpreadsheetMaxSelection(spreadsheet));
		}
	}

	public void syncAutoFilterStatus() {
		final Worksheet worksheet = spreadsheet.getSelectedSheet();
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
		clearFilter.setDisabled(!appliedFilter);
		reapplyFilter.setDisabled(!appliedFilter);
	}
	
	public void onClick$filter() {
		Rect rect = spreadsheet.getSelection();
		searchInfo.clear();
		Range range = searchAutoFilterRange(Integer.valueOf(rect.getTop()), Integer.valueOf(rect.getLeft()));
		if (range != null) {
			range.autoFilter();
		}
	}
	
	public void onClick$clearFilter() {
		Ranges.range(spreadsheet.getSelectedSheet()).showAllData();
	}
	
	public void onClick$reapplyFilter() {
		Ranges.range(spreadsheet.getSelectedSheet()).applyFilter();
	}

	public void onFilter$spreadsheet(CellMouseEvent event) {
		Worksheet sheet = spreadsheet.getSelectedSheet();
		AutoFilter autoFilter = sheet.getAutoFilter();
		CellRangeAddress cellRangeAddr = autoFilter.getRangeAddress();
		int left = cellRangeAddr.getFirstColumn();
		int top = cellRangeAddr.getFirstRow();
		int right = cellRangeAddr.getLastColumn();
		Range range = 
			Ranges.range(spreadsheet.getSelectedSheet(), cellRangeAddr.getFirstRow(), left, cellRangeAddr.getLastRow(), right);
		int fieldIdx = 0 ;
		for (int i = left; i <= right; i++) {
			fieldIdx += 1;
			Rect rect = getMergedRect(Utils.getCell(sheet, top, i));
			if (rect != null && rect.getRight() != -1) {
				i = rect.getRight();
			}
			if (i == event.getColumn()) {
				break;
			}
		}
		openAutoFilterDialog(new Object[]{fieldIdx, event.getColumn(), range});
	}
	
	/**
	 * Returns whether cell is inside range or not
	 */
	private boolean insideRange(Range src, Cell cell) {		
		int row = cell.getRowIndex();
		int col = cell.getColumnIndex();
		return row >= src.getRow() && row <= src.getLastRow() && 
			col >= src.getColumn() && col <= src.getLastColumn();
	}
	
	private Rect getMergedRect(Cell cell) {
		if (cell == null)
			return null;
		org.zkoss.poi.ss.usermodel.Sheet sheet = cell.getSheet();
		int row = cell.getRowIndex();
		int col = cell.getColumnIndex();
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
			org.zkoss.poi.ss.util.CellRangeAddress ref = sheet.getMergedRegion(i);
			int top = ref.getFirstRow();
	        int btm = ref.getLastRow();
	        int left = ref.getFirstColumn();
	        int right = ref.getLastColumn();
	        
	        if (row >= ref.getFirstRow() && row <= ref.getLastRow() &&
	        	col >= ref.getFirstColumn() && col <= ref.getLastColumn())
	        	return new Rect(left, top, right, btm);
		}
		return null;
	}
	
	ArrayDeque<Range> searchInfo = new ArrayDeque<Range>();
	private Range getCurrentAutoFilterRange() {
		return searchInfo.size() > 0 ? searchInfo.getFirst() : null;
	}
	private Range searchAutoFilterRange(Integer row, Integer col) {
		Range rng = getCurrentAutoFilterRange();
		Worksheet sheet = spreadsheet.getSelectedSheet();
		if (rng == null) {
			for (int i = row - 1; i <= row + 1; i++) {
				for (int j = col - 1; j <= col + 1 ; j++) {
					Cell c = i != row && j != col ? Utils.getCell(sheet, i, j) : null;
					getMergedRect(c);
					if (c != null && c.getCellType() != Cell.CELL_TYPE_BLANK) {
						boolean verticalOrient = i > row; /* false means top; true means bottom */
						boolean horizontalOrient = j > col; /* fasle means left; true means right */
						rng = Ranges.range(sheet, verticalOrient ? row : i, horizontalOrient ? col : j, 
								verticalOrient ? i : row, horizontalOrient ? j : col);
						searchInfo.push(rng);
						searchAutoFilterRange(null, null);
					}
				}
			}
		} else {
			int top = rng.getRow();
			int btm = rng.getLastRow();
			int left = rng.getColumn();
			int right = rng.getLastColumn();
			
			for (int i = top - 1; i <= btm + 1; i++) {
				for (int j = left - 1; j <= right + 1 ; j++) {
					if (i < 0 || j < 0)
						continue;
					Cell c = Utils.getCell(sheet, i, j);
					if (c != null && !insideRange(getCurrentAutoFilterRange(), c)) {
						int cellType = c.getCellType();
						Rect mergeRect = null;
						if (cellType == Cell.CELL_TYPE_BLANK) {
							mergeRect = getMergedRect(c);
							if (mergeRect == null)
								continue;
						}
						if (mergeRect != null) {
							int t = Math.min(Math.min(i, top), mergeRect.getTop());
							int l = Math.min(Math.min(j, left), mergeRect.getLeft());
							int b = Math.max(Math.max(i, btm), mergeRect.getBottom());
							int r = Math.max(Math.max(j, right), mergeRect.getRight());
							rng = Ranges.range(sheet, t, l, b, r);
						} else 
							rng = Ranges.range(sheet, 
								Math.min(i, top), Math.min(j, left), Math.max(i, btm), Math.max(j, right));
						searchInfo.pop();
						searchInfo.push(rng);
						searchAutoFilterRange(null, null);
					}
				}
			}
		}
		return searchInfo.size() > 0 ? searchInfo.getFirst() : null;
	}
	
	public void onClick$insertHyperlinkBtn() {
		openHyperlinkDialog();
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}
	
	public void onClick$cutBtn() {
		getDesktopWorkbenchContext().getWorkbookCtrl().cutSelection();
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}
	
	public void onClick$copyBtn() {
		getDesktopWorkbenchContext().getWorkbookCtrl().copySelection();
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}

	public void onClick$closeBtn() {
		getDesktopWorkbenchContext().getWorkbookCtrl().close();
		getDesktopWorkbenchContext().fireWorkbookChanged();
	}
	
	protected DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return Zssapp.getDesktopWorkbenchContext(self);
	}
	
	public void redrawSheetTabbox() {
		sheets.redraw();
	}
	
	public void setSelectedSheet(String name) {
		spreadsheet.setSelectedSheet(name);
		gridlinesCheckbox.setChecked(spreadsheet.getSelectedSheet().isDisplayGridlines());
	}

	public void updateGridlinesCheckbox() {
		gridlinesCheckbox.setChecked(spreadsheet.getSelectedSheet().isDisplayGridlines());
	}
	
	// SECTION Spreadsheet Event Handler
	void doMouseEvent(CellMouseEvent event) {
		event_x = event.getClientx();
		event_y = event.getClienty();

		cellMenupopup.open(event_x + 5, event.getClienty());
	}

	void doHeaderMouseEvent(HeaderMouseEvent event) {
		if (HeaderEvent.TOP_HEADER == event.getType()) {
			int col = event.getIndex();
			columnHeaderMenupopup.open(event.getClientx(), event.getClienty());
		} else if ((HeaderEvent.LEFT_HEADER == event.getType())) {
			int row = event.getIndex();
			rowHeaderMenupopup.open(event.getClientx(), event.getClienty());
		} else {
			return;
		}
	}

	void doFocusedEvent(final CellEvent event) {
		final Book book = spreadsheet.getBook();
		if (book == null) {
			return;
		}
		// SECTION WORK1 FocusedEvent
		try {
			Worksheet sheet = event.getSheet();
			lastRow = event.getRow();
			lastCol = event.getColumn();

			Cell cell = Utils.getOrCreateCell(sheet, lastRow, lastCol);

			// reading the text from cell to formula bar
			CellReference cr = new CellReference(lastRow, lastCol, false, false);
			String posStr = cr.formatAsString();
			lbpos.setLabel(posStr);
			focusPosition.setRawValue(posStr);
			focusPosition.setSelectedItem(lbpos);

			// read format from cell and assign it to toolbar
			// merge cell
			isMergeCell = isMergedCell(event.getRow(), event.getColumn(), event.getRow(), event.getColumn());
			mergeCellBtn.setSclass(isMergeCell ? "clicked" : null);

			getCellStyleContext().doTargetChange(new SSRectCellStyle(cell, spreadsheet));

		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	private boolean isMergedCell(int tRow, int lCol, int bRow, int rCol) {
		MergeMatrixHelper mmhelper = spreadsheet.getMergeMatrixHelper(spreadsheet.getSelectedSheet());
		for (final Iterator iter = mmhelper.getRanges().iterator(); iter
				.hasNext();) {
			MergedRect block = (MergedRect) iter.next();
			int bt = block.getTop();
			int bl = block.getLeft();
			int bb = block.getBottom();
			int br = block.getRight();
			if (lCol <= bl && tRow <= bt && rCol >= br && bRow >= bb) {
				return true;
			}
		}
		return false;
	}

	public void doSelectionEvent(CellSelectionEvent event) {
		isMergeCell = isMergedCell(
				spreadsheet.getSelection().getTop(), 
				spreadsheet.getSelection().getLeft(), 
				spreadsheet.getSelection().getBottom(),
				spreadsheet.getSelection().getRight());

		mergeCellBtn.setSclass(isMergeCell ? "clicked" : null);
	}

	// Enter key
	public void onOK() {
		EditHelper.doPaste(spreadsheet);
		EditHelper.clearCutOrCopy(spreadsheet);
	}
	
	// ESC key
	public void onCancel() {
		EditHelper.clearCutOrCopy(spreadsheet);
	}

	// SECTION CtrlKeys
	public void onSSCtrlKeys(ForwardEvent event) {
		Event orig = event.getOrigin();
		while (orig instanceof ForwardEvent) {
			orig = ((ForwardEvent) orig).getOrigin();
		}
		if (orig instanceof KeyEvent) {
			_onSSCtrlKeys((KeyEvent) orig);
		}
	}

	public void _onSSCtrlKeys(KeyEvent event) {
		char c = (char) event.getKeyCode();
		try {
			if (46 == event.getKeyCode()) {// delete
				if (event.isCtrlKey())
					CellHelper.clearStyle(spreadsheet, 
							SheetHelper.getSpreadsheetMaxSelection(spreadsheet));
				else
					CellHelper.clearContent(spreadsheet,
							SheetHelper.getSpreadsheetMaxSelection(spreadsheet));
				return;
			}

			if (false == event.isCtrlKey())
				return;

			switch (c) {
			case 'X':
				EditHelper.doCut(spreadsheet);
				break;
			case 'C':
				EditHelper.doCopy(spreadsheet);
				break;
			case 'V':
				EditHelper.doPaste(spreadsheet);
				break;
			case 'D':
				CellHelper.clearContent(spreadsheet,
						SheetHelper.getSpreadsheetMaxSelection(spreadsheet));
				break;
			case 'E':// for testing
				spreadsheet.setCellFocus(new Position(2, 2));
				spreadsheet.setSelection(new Rect(2, 2, 4, 4));
				break;
			case 'S':
				//TODO: check permission from WorkbookCtrl
				if (FileHelper.hasSavePermission()) {
					//TODO: refactor duplicate save logic
					DesktopWorkbenchContext workbench = getDesktopWorkbenchContext();
					if (workbench.getWorkbookCtrl().hasFileExtentionName()) {
						workbench.getWorkbookCtrl().save();
						workbench.fireWorkbookSaved();
					} else
						workbench.getWorkbenchCtrl().openSaveFileDialog();
				}
				break;
			// TODO undo
			/*
			 * case 'Z': spreadsheet.undo(); break;
			 */
			// TODO redo
			/*
			 * case 'Y': spreadsheet.redo(); break;
			 */
			case 'O':
				openOpenFileDialog();
				break;
			case 'F':
				//TODO: 
				break;
			case 'B':
				getCellStyleContext().modifyStyle(new StyleModification(){
					public void modify(org.zkoss.zss.app.zul.ctrl.CellStyle style, CellStyleContextEvent candidteEvt) {
						candidteEvt.setExecutor(MainWindowCtrl.this);
						style.setBold(!style.isBold());
					}
				});
				break;
			case 'I':
				getCellStyleContext().modifyStyle(new StyleModification(){
					public void modify(org.zkoss.zss.app.zul.ctrl.CellStyle style, CellStyleContextEvent candidteEvt) {
						candidteEvt.setExecutor(MainWindowCtrl.this);
						style.setItalic(!style.isItalic());
					}
				});
				break;
			case 'U':
				getCellStyleContext().modifyStyle(new StyleModification(){
					public void modify(org.zkoss.zss.app.zul.ctrl.CellStyle style, CellStyleContextEvent candidteEvt) {
						candidteEvt.setExecutor(MainWindowCtrl.this);
						boolean isUnderline = style.getUnderline() ==  org.zkoss.zss.app.zul.ctrl.CellStyle.UNDERLINE_SINGLE;
						style.setUnderline(isUnderline ? 
								org.zkoss.zss.app.zul.ctrl.CellStyle.UNDERLINE_NONE : org.zkoss.zss.app.zul.ctrl.CellStyle.UNDERLINE_SINGLE);
					}
				});
				break;
			default:
				return;
			}

		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * @return
	 */
	private DesktopCellStyleContext getCellStyleContext() {
		return Zssapp.getDesktopCellStyleContext(self);
	}

	public void onExportFile(ForwardEvent event) {
		throw new UiException("export file not implement yet");
//		Listbox fle_files = (Listbox) Path
//				.getComponent("//p1/mainWin/fileExportWin/fle_files");
//
//		String filename = fle_files.getSelectedItem().getLabel();
//		exportFile(filename);
//
//		Window fileExportWin = (Window) Path
//				.getComponent("//p1/mainWin/fileExportWin");
//		fileExportWin.setVisible(false);
	}

	public void exportFile(String filename) {// current or other
												// files(stack_level=0)
	// TODO
		throw new UiException("export file not implmented yet");
		/*
		 * // current editing file System.out.println("exportFile filename: " +
		 * filename + ", spreadsheet source: " + spreadsheet.getSrc()); if
		 * (filename.equals(spreadsheet.getSrc())) {
		 * System.out.println("exportFile if: "); ByteArrayOutputStream
		 * baoStream = new ByteArrayOutputStream(); ExcelExporter exporter = new
		 * ExcelExporter(); exporter.exports(spreadsheet.getBook(), baoStream);
		 * byte[] bin = baoStream.toByteArray(); Filedownload.save(new
		 * ByteArrayInputStream(bin), "application/vnd.ms-excel", filename); }
		 * else {// other files System.out.println("exportFile else"); HashMap
		 * hm = fileh.readMetafile(); Object objs[] = (Object[])
		 * hm.get(filename); if (objs != null) { String hashFilename =
		 * fileh.xlsDir + ((String) objs[1]);
		 * System.out.println("hashFilename: " + hashFilename); FileInputStream
		 * iStream = null; try { iStream = new FileInputStream(hashFilename); }
		 * catch (FileNotFoundException e) { e.printStackTrace(); }
		 * Filedownload.save(iStream, "application/vnd.ms-excel", filename); } }
		 */}

	public void onPrint(ForwardEvent event) {
		throw new UiException("print not implement yet");
//		String printKey = "" + System.currentTimeMillis();
//		Session session = Executions.getCurrent().getDesktop().getSession();
//		session.setAttribute("zssFromHi" + printKey, spreadsheet);
//
//		Window win = (Window) mainWin.getFellow("menuPrintWin");
//		Button printBtn = (Button) win.getFellow("printBtn");
//		printBtn.setHref("print.zul?printKey=" + printKey);
//		printBtn.setTooltip("print.zul?printKey=" + printKey);
//
//		win.setPosition("parent");
//		win.setTop("100px");
//		win.setLeft("100px");
//		win.doPopup();
	}

	public void onRevision(ForwardEvent event) {
		throw new UiException("revision not implement yet");
//		reloadRevisionMenu();
//		Window win = (Window) mainWin.getFellow("revisionWin");
//		win.setPosition("parent");
//		win.setLeft("250px");
//		win.setTop("250px");
//		win.doPopup();
	}

	// SECTION EDIT MENU
	public void onEditUndo(ForwardEvent event) {
		throw new UiException("Undo not implmented yet");
		// spreadsheet.undo();
	}

	public void onEditRedo(ForwardEvent event) {
		throw new UiException("Redo not implemented yet");
		// spreadsheet.redo();
	}

	//TODO: move these method to compositive component
	public void onEditCut(ForwardEvent event) {
		EditHelper.doCut(spreadsheet);
	}

	public void onEditCopy(ForwardEvent event) {
		EditHelper.doCopy(spreadsheet);
	}

	public void onEditPaste(ForwardEvent event) {
		EditHelper.doPaste(spreadsheet);
	}

	public void onDeleteRows(ForwardEvent event) {
		try {
			Worksheet sheet = spreadsheet.getSelectedSheet();
			Rect rect = spreadsheet.getSelection();
			int top = rect.getTop();
			int left = rect.getLeft();
			int bottom = rect.getBottom();
			if (top == 0 && bottom == spreadsheet.getMaxrows() - 1) {
				try {
					Messagebox.show("cannot delete all Rows");
					return;
				} catch (InterruptedException e) {
				}
			}
			// TODO undo/redo
			// spreadsheet.pushDeleteRowColState(-1, top, -1, bottom);
			Utils.deleteRows(sheet, top, bottom);
			spreadsheet.focus();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void onDeleteColumns(ForwardEvent event) {
		try {
			Worksheet sheet = spreadsheet.getSelectedSheet();
			Rect rect = spreadsheet.getSelection();
			int left = rect.getLeft();
			int right = rect.getRight();
			if (left == 0 && right == spreadsheet.getMaxcolumns() - 1) {
				try {
					Messagebox.show("cannot delete all Columns");
					return;
				} catch (InterruptedException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
			// TODO undo/redo
			// spreadsheet.pushDeleteRowColState(left, -1, right, -1);
			Utils.deleteColumns(sheet, left, right);
			spreadsheet.focus();

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private boolean onViewFormulaBar() {
		formulaBar.setVisible(!formulaBar.isVisible());
		return formulaBar.isVisible();
	}

	// SECTION FORMAT MENU
	public void onFormatNumber(ForwardEvent event) {
		Window win = (Window) mainWin.getFellow("formatNumberWin");
		win.setPosition("parent");
		win.setLeft("170px");
		win.setTop("24px");
		win.doPopup();// Modal();
	}

	// SECTION INSERT MENU
	public void onInsertFormula(ForwardEvent event) {
		Window win = (Window) mainWin.getFellow("formulaCategory");
		win.doHighlighted();
	}

	public void onInsertRows(ForwardEvent event) {
		try {
			final Rect rect = spreadsheet.getSelection();
			final int top = rect.getTop();
			final int bottom = rect.getBottom();
			// TODO undo/redo
			// spreadsheet.pushInsertRowColState(-1, top, -1, bottom);
			Utils.insertRows(spreadsheet.getSelectedSheet(), top, bottom);
			spreadsheet.focus();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void onInsertColumns(ForwardEvent event) {
		try {
			Rect rect = spreadsheet.getSelection();
			int left = rect.getLeft();
			int right = rect.getRight();

			// TODO undo/redo
			// spreadsheet.pushInsertRowColState(left, -1, right, -1);
			Utils.insertColumns(spreadsheet.getSelectedSheet(), left, right);
			spreadsheet.focus();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void onInsertSheet() {
		final Book book = spreadsheet.getBook();
		if (book == null) {
			return;
		}
		int sheetCount = book.getNumberOfSheets();
		Worksheet addedSheet = (Worksheet)spreadsheet.getBook().createSheet("sheet " + (sheetCount + 1));
		sheets.addSheet(addedSheet.getSheetName());
	}

	// SECTION HELP MENU
	public void onHelpCheatsheet(ForwardEvent event) {
		Window win = (Window) mainWin.getFellow("cheatsheet");
		win.setPosition("parent");
		win.setLeft("327px");
		win.setTop("124px");
		win.doPopup();
	}

	public void onFormulaListOK() {
		Combobox formulaList = (Combobox) Path.getComponent("//p1/mainWin/formulaList");

		int left = spreadsheet.getSelection().getLeft();
		int top = spreadsheet.getSelection().getTop();
		String formula = formulaList.getSelectedItem().getLabel();
		if (formula != null) {
			final Worksheet sheet = spreadsheet.getSelectedSheet();
			Cell cell = Utils.getOrCreateCell(sheet, top, left);
			cell.setCellFormula(formula + "()");
		}
		formulaList.setText("");
	}

	public void onMergeCellClick(ForwardEvent event) {
		try {
			// isMergeCell should read from the cell
			// and search all over the cell if any one is merged then unmerge
			// all
			isMergeCell = !isMergeCell;
			if (isMergeCell) {
				mergeCellBtn.setClass("clicked");
			}

			if (isMergeCell) {
				// TODO: undo/redo
				// spreadsheet.pushCellState();
				Rect sel = spreadsheet.getSelection();
				if (sel.getLeft() == sel.getRight() && sel.getTop() == sel.getBottom()) {
					mergeCellBtn.setClass("toolIcon");
				} else {
					Utils.mergeCells(spreadsheet.getSelectedSheet(), sel.getTop(), sel.getLeft(), sel.getBottom(), sel.getRight(), false);
				}
				spreadsheet.focus();
			} else {
				// TODO: undo/redo
				// spreadsheet.pushCellState();
				Rect sel = spreadsheet.getSelection();
				Utils.unmergeCells(spreadsheet.getSelectedSheet(),
						sel.getTop(), sel.getLeft(), sel.getBottom(), sel
								.getRight());
				spreadsheet.focus();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}



	public void onRange(ForwardEvent event) {
		rangeh.dispatcher((String) event.getData());
	}



	public void onClick$insertPieChart(Event evt) {
		// TODO remove me, insert PieChart
		throw new UiException("insert pie chart is not implmented yet");
		/*
		 * Component cmp = execution.createComponents("piechart.zul", mainWin,
		 * null); if (cmp != null) { System.out.println("create cmp"); } Chart
		 * chart = (Chart)cmp.getFellow("mychart"); ChartWidget wgt = new
		 * ChartWidget(); wgt.setChart(chart);
		 * 
		 * int col = spreadsheet.getSelection().getLeft(); int row =
		 * spreadsheet.getSelection().getTop(); wgt.setRow(row);
		 * wgt.setColumn(col); SpreadsheetCtrl ctrl = (SpreadsheetCtrl)
		 * spreadsheet.getExtraCtrl(); ctrl.addWidget(wgt);
		 */}

	public void onInsertChart(ForwardEvent event) {
		// TODO remove me, insert chart
		throw new UiException("insert chart is not implmented yet");
		/*
		 * final int left = spreadsheet.getSelection().getLeft(); final int
		 * right = spreadsheet.getSelection().getRight(); final int top =
		 * spreadsheet.getSelection().getTop(); final int bottom =
		 * spreadsheet.getSelection().getBottom(); Sheet sheet =
		 * spreadsheet.getSelectedSheet();
		 * 
		 * final String modelType=(String)event.getData();
		 * 
		 * //BarModel model2=new SimpleBarModel(); StringBuffer strBuf = new
		 * StringBuffer();
		 * strBuf.append("<window title=\"test\" width=\"420px\" id=\"chartWin"
		 * +(chartKey++)+"\" border=\"normal\">");strBuf.append(
		 * "<chart id=\"myChart\" width=\"400px\" height=\"200px\" type=\""
		 * +modelType+"\" threeD=\"true\" fgAlpha=\"128\">");
		 * strBuf.append("<zscript>"); p(modelType);
		 * if(modelType.equals("pie")){
		 * strBuf.append("myChart.setModel(new SimplePieModel());"); }
		 * if(modelType.equals("bar")){
		 * strBuf.append("myChart.setModel(new SimpleCategoryModel());"); }
		 * strBuf.append("</zscript>"); strBuf.append("</chart>");
		 * strBuf.append("</window>"); Window win= (Window)
		 * Executions.createComponentsDirectly(strBuf.toString(), null, mainWin,
		 * null); Chart myChart = (Chart) win.getFellow("myChart"); ChartModel
		 * model=myChart.getModel(); for(int row=top; row<=bottom; row++){ Cell
		 * cellName=sheet.getCell(row, left); String name; if(cellName==null ||
		 * cellName.getText()==null) name=""; else name=cellName.getText();
		 * for(int col=left+1; col<=right; col++){ Cell
		 * cellValue=sheet.getCell(row, col); double value; if(cellValue==null
		 * || cellValue.getResult()==null) value=0; else{ try{
		 * value=((Double)cellValue.getResult()).doubleValue(); }catch(Exception
		 * e){ value=0; }
		 * 
		 * //p(name); //p(""+value); if(modelType.equals("pie")){
		 * ((PieModel)model).setValue((Comparable)name,new Double(value)); }
		 * if(modelType.equals("bar")){
		 * ((CategoryModel)model).setValue((Comparable)name, (Comparable)new
		 * Long(col-left), new Double(value)); } } } }
		 * if(modelType.equals("pie")) myChart.setModel((PieModel)model);
		 * if(modelType.equals("bar")) myChart.setModel((CategoryModel)model);
		 * 
		 * win.doOverlapped(); //win.setSizable(true); win.setClosable(true);
		 * //win.setAttribute(SpreadsheetCtrl.CHILD_PASSING_KEY,"");
		 * //spreadsheet.appendChild(win); mainWin.appendChild(win);
		 * 
		 * win.setLeft("100px"); win.setTop("200px"); win.doOverlapped();
		 * //spreadsheet.smartUpdate("appendWidget",win.getUuid());
		 * 
		 * 
		 * 
		 * 
		 * //Add EventListener for updating chart
		 * myChart.addEventListener(Events.ON_STOP_EDITING, new EventListener()
		 * { public void onEvent(Event event) throws Exception { int
		 * modRow=((CellEvent)event).getRow(); int
		 * modCol=((CellEvent)event).getColumn(); if(left<=modCol &&
		 * modCol<=right && top<=modRow && modRow<=bottom){
		 * //p("ChartValue Editing:"+(String)event.getData());
		 * //p("row: "+top+" "+bottom); //p("col: "+left+" "+right);
		 * //p("event:"+modelType); Sheet sheet=((CellEvent)event).getSheet();
		 * Chart myChart=(Chart)event.getTarget(); ChartModel model =
		 * myChart.getModel(); if(modelType.equals("pie"))
		 * ((PieModel)model).clear(); if(modelType.equals("bar"))
		 * ((CategoryModel)model).clear();
		 * 
		 * for(int row=top; row<=bottom; row++){ Cell
		 * cellName=sheet.getCell(row, left); String name; if(cellName==null ||
		 * cellName.getText()==null) name=""; else name=cellName.getText();
		 * 
		 * for(int col=left+1; col<=right; col++){ Cell
		 * cellValue=sheet.getCell(row, col); double value; if(cellValue==null
		 * || cellValue.getResult()==null) value=0; else{ try{
		 * value=((Double)cellValue.getResult()).doubleValue(); }catch(Exception
		 * e){ value=0; }
		 * 
		 * //p(name); //p(""+value); if(modelType.equals("pie")){
		 * ((PieModel)model).setValue((Comparable)name,new Double(value)); }
		 * if(modelType.equals("bar")){
		 * ((CategoryModel)model).setValue((Comparable)name, (Comparable)new
		 * Long(col-left), new Double(value)); } } } }
		 * if(modelType.equals("pie")) myChart.setModel((PieModel)model);
		 * if(modelType.equals("bar")) myChart.setModel((CategoryModel)model);
		 * 
		 * }
		 * 
		 * } });
		 * 
		 * win.addEventListener(org.zkoss.zk.ui.event.Events.ON_CLICK, new
		 * EventListener(){
		 * 
		 * public void onEvent(Event event) throws Exception { Rect rect=new
		 * Rect(); rect.setBottom(bottom); rect.setTop(top);
		 * rect.setRight(right); rect.setLeft(left);
		 * spreadsheet.setSelection(rect); } });
		 */
	}

	public void reloadRevisionMenu() {
		throw new UiException("reloadRevision not implement yet");
//		try {
//			// read the metafile
//			BufferedReader bReader;
//			bReader = new BufferedReader(new FileReader(fileh.xlsDir
//					+ "metaFile"));
//			String line = null;
//			Stack stack = new Stack();
//
//			String timeStr, filename, hashFilename;
//			while (true) {
//				timeStr = bReader.readLine();
//				if (timeStr == null) {
//					break;
//				}
//				filename = bReader.readLine();
//				if (filename == null) {
//					System.out.println("Warning: filename cannot read from metaFile");
//					break;
//				}
//				hashFilename = bReader.readLine();
//				if (hashFilename == null) {
//					System.out.println("Warning: hashFilename cannot read from metaFile");
//					break;
//				}
//				if (spreadsheet.getSrc().equals(filename)) {
//					stack.add(timeStr);
//					stack.add(filename);
//					stack.add(hashFilename);
//				}
//			}
//			bReader.close();
//
//			// put read data to Menu
//			org.zkoss.zul.Row newRow;
//			Rows revisionRows = (Rows) mainWin.getFellow("revisionWin")
//					.getFellow("revisionRows");
//			List rowsChildList = revisionRows.getChildren();
//			while (!rowsChildList.isEmpty())
//				revisionRows.removeChild((Component) rowsChildList.get(0));
//			while (!stack.isEmpty()) {
//
//				hashFilename = (String) stack.pop();
//				filename = (String) stack.pop();
//				timeStr = (String) stack.pop();
//				String dateStr = new Date(Long.parseLong(timeStr)).toString();
//
//				newRow = new org.zkoss.zul.Row();
//				Radio radio = new Radio();
//				radio.setAttribute("value", hashFilename);
//
//				if (hashFilename.equals(this.hashFilename)) {
//					newRow.appendChild(new Label("current"));
//					newRow.setStyle("background:rgb(250,230,180) none");
//				} else {
//					newRow.appendChild(radio);
//					// p("set hashFilename: "+hashFilename);
//				}
//				newRow.appendChild(new Label(dateStr));
//				newRow.appendChild(new Label("test name"));
//				newRow.appendChild(new Label("no comment"));
//
//				revisionRows.appendChild(newRow);
//				// p(""+radio.getAttribute("value"));
//			}
//			// close the jdbc connection
//
//		} catch (Exception ex) {
//			throw new RuntimeException(ex);
//		}
	}

//	public void onRevisionOK(ForwardEvent event) {
//		String filename = null;
//
//		Rows revisionRows = (Rows) mainWin.getFellow("revisionWin").getFellow(
//				"revisionRows");
//		List rowList = revisionRows.getChildren();
//		for (int i = 0; i < rowList.size(); i++) {
//			Component tmpComponent = ((Component) rowList.get(i))
//					.getFirstChild();
//			if (tmpComponent instanceof Radio
//					&& ((Radio) tmpComponent).isChecked()) {
//				filename = (String) ((Radio) tmpComponent)
//						.getAttribute("value");
//				// p("onRevision: "+filename);
//			}
//		}
//
//		openFileInFS(filename);
//
//		Window revisionWin = (Window) Path
//				.getComponent("//p1/mainWin/revisionWin");
//		revisionWin.setVisible(false);
//		spreadsheet.invalidate();
//		// spreadsheet.notifyRevision();
//	}

	private Connection getDBConnection() {
		try {
			return DriverManager.getConnection(
					"jdbc:mysql://localhost:3306/zss", "root", "rootzk");
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}

	public void onNotImplement(ForwardEvent event) {
		try {
			Messagebox.show("Not implement yet");
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
	}
	
	public void onCheck$gridlinesCheckbox() {
		Worksheet sheet = spreadsheet.getSelectedSheet();
		Utils.getRange(sheet, 0, 0).setDisplayGridlines(!sheet.isDisplayGridlines());
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}
	
	public void onCheck$protectSheet() {
		getDesktopWorkbenchContext().getWorkbookCtrl().protectSheet(protectSheet.isChecked() ? "" : null);
	}

	public void openCustomSortDialog() {
		if (_customSortDialog == null || _customSortDialog.isInvalidated())
			_customSortDialog = (Dialog) Executions.createComponents(Consts._CustomSortDialog_zul, mainWin, CustomSortWindowCtrl.newArg(spreadsheet));
		_customSortDialog.fireOnOpen(null);
	}
	
	public void openHyperlinkDialog() {
		if (_insertHyperlinkDialog == null || _insertHyperlinkDialog.isInvalidated())
			_insertHyperlinkDialog = (Dialog)Executions.createComponents(Consts._InsertHyperlinkDialog_zul, mainWin, Zssapps.newSpreadsheetArg(spreadsheet));
		_insertHyperlinkDialog.fireOnOpen(null);
	}

	public void openInsertFormulaDialog() {
		if (_insertFormulaDialog == null || _insertFormulaDialog.isInvalidated())
			_insertFormulaDialog = (Dialog)Executions.createComponents(Consts._InsertFormulaDialog2_zul, mainWin, null);
		_insertFormulaDialog.fireOnOpen(null);
	}

	public void openPasteSpecialDialog() {
		if (_pasteSpecialDialog == null || _pasteSpecialDialog.isInvalidated())
			_pasteSpecialDialog = (Dialog)Executions.createComponents(Consts._PasteSpecialDialog_zul, mainWin, Zssapps.newSpreadsheetArg(spreadsheet));
		_pasteSpecialDialog.fireOnOpen(null);
	}

	public boolean toggleFormulaBar() {
		return onViewFormulaBar();
	}

	public void openComposeFormulaDialog(FormulaMetaInfo metainfo) {
		if (_composeFormulaDialog == null || _composeFormulaDialog.isInvalidated())
			_composeFormulaDialog = (Dialog) Executions.createComponents(Consts._ComposeFormulaDialog_zul, mainWin, arg);
		_composeFormulaDialog.fireOnOpen(metainfo);
	}

	public void openFormatNumberDialog() {
		if (_formatNumberDialog == null || _formatNumberDialog.isInvalidated())
			_formatNumberDialog = (Dialog) Executions.createComponents(Consts._FormatNumberDialog_zul, mainWin, Zssapps.newSpreadsheetArg(spreadsheet));
		_formatNumberDialog.fireOnOpen(null);	
	}

	public void openSaveFileDialog() {
		if (_saveFileDialog == null || _saveFileDialog.isInvalidated())
			_saveFileDialog = (Dialog) Executions.createComponents(Consts._SaveFile_zul, mainWin, null);
		_saveFileDialog.fireOnOpen(null);
	}

	public void openModifyHeaderSizeDialog(int headerType) {
		Rect seld = spreadsheet.getSelection();
		int prev = -1;
		boolean sameVal = true;
		if (headerType == WorkbookCtrl.HEADER_TYPE_ROW) {
			for (int i = seld.getTop(); i <= seld.getBottom(); i++) {
				if (prev < 0)
					prev = Utils.getRowHeightInPx(spreadsheet.getSelectedSheet(), i);
				else if (prev != Utils.getRowHeightInPx(spreadsheet.getSelectedSheet(), i)) {
					sameVal = false;
					break;
				}
			}
		} else {
			for (int i = seld.getLeft(); i <= seld.getRight(); i++) {
				if (prev < 0)
					prev = Utils.getColumnWidthInPx(spreadsheet.getSelectedSheet(), i);
				else if (prev != Utils.getColumnWidthInPx(spreadsheet.getSelectedSheet(), i)) {
					sameVal = false;
					break;
				}
			}
		}

		if (_headerSizeDialog == null || _headerSizeDialog.isInvalidated())
			_headerSizeDialog = (Dialog) Executions.createComponents(Consts._HeaderSize_zul, mainWin, null);
		_headerSizeDialog.fireOnOpen(
			HeaderSizeCtrl.newArg(
				Integer.valueOf(headerType), sameVal ? prev : null));
	}

	public void openRenameSheetDialog(String originalSheetName) {
		if (_renameSheetDialog == null || _renameSheetDialog.isInvalidated())
			_renameSheetDialog = (Dialog) Executions.createComponents(Consts._RenameDialog_zul, mainWin, null);
		_renameSheetDialog.fireOnOpen(RenameSheetCtrl.newArg(originalSheetName));
	}

	public void openOpenFileDialog() {
		if (_openFileDialog == null || _openFileDialog.isInvalidated())
			_openFileDialog = (Dialog) Executions.createComponents(Consts._OpenFile_zul, mainWin, null);
		_openFileDialog.fireOnOpen(null);
	}

	public void openImportFileDialog() {
		if (_importFileDialog == null || _importFileDialog.isInvalidated())
			_importFileDialog = (Dialog) Executions.createComponents(Consts._ImportFile_zul, mainWin, null);
		_importFileDialog.fireOnOpen(null);
	}
	
	public void openExportPdfDialog() {
		if (!hasZssPdf()) {
			try {
				Messagebox.show("Please download Zss Pdf from ZK");
			} catch (InterruptedException e) {
			}
			return;
		}
		Executions _exportToPdfDialogExecutions;
		if (_exportToPdfDialog == null || _exportToPdfDialog.isInvalidated())
			_exportToPdfDialog = (Dialog) Executions.createComponents(Consts._ExportToPDF_zul, mainWin, Zssapps.newSpreadsheetArg(spreadsheet));
		_exportToPdfDialog.fireOnOpen(null);
	}
	
	public void openAutoFilterDialog(Object param) {
		//TODO: If _autoFilterDialog needs cell has client side uuid, then it could use Popup and specify position
		if (_autoFilterDialog == null || _autoFilterDialog.isInvalidated())
			_autoFilterDialog = (Dialog) Executions.createComponents(Consts._AutoFilter_zul, mainWin, null);
		_autoFilterDialog.fireOnOpen(param);
	}
		
	private static boolean hasZssPdf() {
		String val = Library.getProperty(KEY_PDF);
		if (val == null) {
			boolean hasZssPdf = verifyZssPdf();
			Library.setProperty(KEY_PDF, String.valueOf(hasZssPdf));
			return hasZssPdf;
		} else {
			return Boolean.valueOf(Library.getProperty(KEY_PDF));
		}
	}

	/**
	 * Verify whether has zss pdf export function or not
	 */
	private static boolean verifyZssPdf() {
		try {
			Class.forName("org.zkoss.zss.model.impl.pdf.PdfExporter");
		} catch (ClassNotFoundException ex) {
			return false;
		}
		return true;
	}

	//TODO: mimic openExportPdfDialog()
	@Override
	public void openExportHtmlDialog() {
		if (!hasZssHtml()) {
			try {
				Messagebox.show("Please download Zss Html from ZK");
			} catch (InterruptedException e) {
			}
			return;
		}
		Executions _exportToHtmlDialogExecutions;
		if (_exportToHtmlDialog == null || _exportToHtmlDialog.isInvalidated())
			_exportToHtmlDialog = (Dialog) Executions.createComponents(Consts._ExportToHTML_zul, mainWin, Zssapps.newSpreadsheetArg(spreadsheet));
		_exportToHtmlDialog.fireOnOpen(null);
		
	}

	private static boolean hasZssHtml() {
		String val = Library.getProperty(KEY_HTML);
		if (val == null) {
			boolean hasZssHtml = verifyZssHtml();
			Library.setProperty(KEY_HTML, String.valueOf(hasZssHtml));
			return hasZssHtml;
		} else {
			return Boolean.valueOf(Library.getProperty(KEY_HTML));
		}
	}

	private static boolean verifyZssHtml() {
		try {
			Class.forName("org.zkoss.zss.model.impl.html.HtmlExporter");
		} catch (ClassNotFoundException ex) {
			return false;
		}
		return true;
	}
}