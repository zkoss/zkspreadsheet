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
import java.util.HashMap;
import java.util.Iterator;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Sheet;
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
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.cell.CellHelper;
import org.zkoss.zss.app.cell.EditHelper;
import org.zkoss.zss.app.event.ExportHelper;
import org.zkoss.zss.app.file.FileHelper;
import org.zkoss.zss.app.file.SpreadSheetMetaInfo;
import org.zkoss.zss.app.formula.FormulaMetaInfo;
import org.zkoss.zss.app.sheet.SheetHelper;
import org.zkoss.zss.app.sort.SortSelector;
import org.zkoss.zss.app.zul.CellContext;
import org.zkoss.zss.app.zul.CellMenupopup;
import org.zkoss.zss.app.zul.ColumnHeaderMenupopup;
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
import org.zkoss.zss.app.zul.ctrl.WorkspaceContext;
import org.zkoss.zss.engine.event.SSDataEvent;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Ranges;
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
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.South;
import org.zkoss.zul.Toolbarbutton;
import org.zkoss.zul.Window;

/**
 * @author Peter Kuo
 * @modify kinda lu
 */
public class MainWindowCtrl extends GenericForwardComposer implements WorkbenchCtrl {

	private static final long serialVersionUID = 1;
	static int event_x = 200;
	static int event_y = 200;

	int lastRow = 0;
	int lastCol = 0;
	boolean isFreezeRow = false;
	boolean isFreezeColumn = false;

	Div toolbarMask;
	
	// For fast Icon
	boolean isMergeCell = false;
	boolean isWrapText = false;
	String colorSelectorTarget = "";

	Cell currentEditcell;

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
	
	CellContext cellContext;
	Menu insertImageMenu;
	Menu insertPieChart;

	FormulaEditor formulaEditor;
	Spreadsheet spreadsheet;
	Sheets sheets;
	Combobox focusPosition;
	
	Comboitem lbpos;
	
	/* Toolbar buttons */
	Dropdownbutton pasteDropdownBtn;
	Dropdownbutton sortDropdownBtn;
	Toolbarbutton closeBtn;
	CellStyleCtrlPanel fontCtrlPanel;
	Toolbarbutton exportToPDFBtn;
	Toolbarbutton insertHyperlinkBtn;
	Toolbarbutton mergeCellBtn;
	Toolbarbutton insertChartBtn;
	Toolbarbutton cutBtn;
	Toolbarbutton copyBtn;
	Toolbarbutton insertFormulaBtn;
	Checkbox gridlinesCheckbox;
	CellMenupopup cellMenupopup;
	
	South formulaBar;
	Borderlayout topToolbars;
	
	/*Dialog*/
	Window insertFormulaDialog;
	Window insertHyperlinkDialog;

	public Window getMainWindow() {
		return mainWin;
	}

	public Spreadsheet getSpreadsheet() {
		return spreadsheet;
	}

	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		initZssappComponents();
		
		init();
		rangeh = new RangeHelper(spreadsheet);
		DesktopWorkbenchContext.getInstance(desktop).fireWorkbookOpen(spreadsheet.getSelectedSheet() != null);
	}
	
	//TODO: remove this mechanism
	private void initZssappComponents() {
		Zssapps.bindSpreadsheet(spreadsheet, this);
	}
	
	public void init() {
		spreadsheet.setSrcName("Untitled");
		final DesktopWorkbenchContext workbenchContext = DesktopWorkbenchContext.getInstance(desktop);
		workbenchContext.setWorkbenchCtrl(this);
		workbenchContext.doTargetChange(new SSWorkbookCtrl(spreadsheet.getBook(), spreadsheet));

		workbenchContext.addEventListener(Consts.ON_SHEET_REFRESH, new EventListener(){
			public void onEvent(Event event) throws Exception {
				sheets.redraw();
			}
		});
		workbenchContext.addEventListener(Consts.ON_WORKBOOK_OPEN, new EventListener() {
			public void onEvent(Event event) throws Exception {
				Boolean isOpen = (Boolean)event.getData();
				toolbarMask.setVisible(!isOpen);
				closeBtn.setVisible(isOpen);

				if (isOpen) {
					gridlinesCheckbox.setChecked(spreadsheet.getSelectedSheet().isDisplayGridlines());
					sheets.redraw();
				} else if (!isOpen)
					sheets.clear();
				
				//TODO: provide clip board interface, to allow save cut, copy, high light info
				//use set setHighlight null can cancel selection, but need to re-store selection when select same sheet again
				spreadsheet.setHighlight(null);
			}
		});
		workbenchContext.addEventListener(Consts.ON_SHEET_CHANGED, new EventListener() {
			public void onEvent(Event event) throws Exception {
				gridlinesCheckbox.setChecked(spreadsheet.getSelectedSheet().isDisplayGridlines());
			}
		});
		workbenchContext.addEventListener(Consts.ON_SHEET_CONTENTS_CHANGED,  new EventListener(){
			public void onEvent(Event event) throws Exception {
				onContentsChanged();
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
		
		WorkspaceContext workspaceContext = WorkspaceContext.getInstance(desktop);
		workspaceContext.addEventListener(Consts.ON_RESOURCE_OPEN, new EventListener() {
			public void onEvent(Event event) throws Exception {
				FileHelper.openSpreadsheet(spreadsheet, 
						(SpreadSheetMetaInfo) event.getData());
				workbenchContext.fireWorkbookOpen(true);
			}
		});
		workspaceContext.addEventListener(Consts.ON_RESOURCE_OPEN_NEW, new EventListener() {
			public void onEvent(Event event) throws Exception {
				FileHelper.openNewSpreadsheet(spreadsheet);
				workbenchContext.fireWorkbookOpen(true);
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
		
		spreadsheet.getBook().subscribe(new EventListener() {
			public void onEvent(Event event) throws Exception {
				String evtName = event.getName();
				if (evtName == SSDataEvent.ON_CONTENTS_CHANGE) {
					onContentsChanged();
				}
			}
		});
	}
	private void onContentsChanged() {
		Sheet seldSheet = spreadsheet.getSelectedSheet();
		Rect seld =  spreadsheet.getSelection();
		int row = seld.getTop();
		int col = seld.getLeft();
		Cell cell = Utils.getCell(seldSheet, row, col);
		if (cell != null) {
			DesktopCellStyleContext.getInstance(desktop).doTargetChange(
				new SSRectCellStyle(cell, 
						spreadsheet) );
			
			formulaEditor.setText(Ranges.range(seldSheet, row, col).getEditText());
		}
	}
	public void onClick$exportToPDFBtn() {
		ExportHelper.doExportToPDF(spreadsheet);
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}
	
	public void onClick$pasteDropdownBtn() {
		getDesktopWorkbenchContext().getWorkbookCtrl().pasteSelection();
	}
	public void onDropdown$pasteDropdownBtn() {
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}
	public void onPasteSelector(ForwardEvent event) {
		EditHelper.onPasteEventHandler(spreadsheet, (String)event.getData());
	}
	public void onDropdown$sortDropdownBtn() {
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}
	public void onClick$insertFormulaBtn() {
		openInsertFormulaDialog();
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}
	public void onSortSelector(ForwardEvent event) {
		//TODO: replace forward event
		String param = (String) event.getData();
		if (param == null)
			return;
		if (param.equals(Labels.getLabel("sort.custom"))) {
			CellHelper.createCustomSortDialog(spreadsheet, mainWin);
		} else {
			if (SortSelector.getSortOrder(param))
				CellHelper.sortAscending(spreadsheet.getSelectedSheet(), SheetHelper.getSpreadsheetMaxSelection(spreadsheet));
			else
				CellHelper.sortDescending(spreadsheet.getSelectedSheet(), SheetHelper.getSpreadsheetMaxSelection(spreadsheet));
		}
	}
	
	public void onClick$insertHyperlinkBtn() {
		openHyperlinkDialog();
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}
	
	public void onClick$cutBtn() {
		getDesktopWorkbenchContext().getWorkbookCtrl().cutSelection();
	}
	
	public void onClick$copyBtn() {
		getDesktopWorkbenchContext().getWorkbookCtrl().copySelection();
	}

	public void onClick$closeBtn() {
		spreadsheet.setSrc(null);
		getDesktopWorkbenchContext().fireWorkbookOpen(false);
	}
	
	protected DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return DesktopWorkbenchContext.getInstance(desktop);
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
			Sheet sheet = event.getSheet();
			lastRow = event.getRow();
			lastCol = event.getColumn();

			Cell cell = Utils.getCell(sheet, lastRow, lastCol);

			// reading the text from cell to formula bar
			CellReference cr = new CellReference(lastRow, lastCol, false, false);
			String posStr = cr.formatAsString();
			lbpos.setLabel(posStr);
			focusPosition.setRawValue(posStr);
			focusPosition.setSelectedItem(lbpos);

			// read format from cell and assign it to toolbar
			if (cell != null) {
				CellStyle cs = cell.getCellStyle();
				if (cs != null) {
					// merge cell
					isMergeCell = isMergedCell(event.getRow(), event.getColumn(), event.getRow(), event.getColumn());
					mergeCellBtn.setSclass(isMergeCell ? "clicked" : null);

					DesktopCellStyleContext.getInstance(desktop).doTargetChange(
							new SSRectCellStyle(cell, 
									spreadsheet) );
				}
			}

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
				//TODO: not implement yet
				FileHelper.saveSpreadsheet(spreadsheet);
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
				Executions.createComponents(Consts._FileListOpen_zul, mainWin, Zssapps.newSpreadsheetArg(spreadsheet));
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
		return DesktopCellStyleContext.getInstance(Executions.getCurrent().getDesktop());
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
			Sheet sheet = spreadsheet.getSelectedSheet();
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
			Sheet sheet = spreadsheet.getSelectedSheet();
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

	//TODO: Do not use hard code
	public void onViewFormulaBar() {
		if (formulaBar.getHeight() != "0px") {
			formulaBar.setSize("0px");
			topToolbars.setHeight("58px");
		} else {
			topToolbars.setHeight("85px");
			formulaBar.setSize("23px");
		}
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
		Sheet addedSheet = spreadsheet.getBook().createSheet("sheet " + (sheetCount + 1));
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
			final Sheet sheet = spreadsheet.getSelectedSheet();
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
				if (sel.getLeft() - sel.getRight() == 0) {
					mergeCellBtn.setClass("toolIcon");
				} else {
					// TODO: true mean merge horizontal only.(UI cannot handle
					// merge vertically yet)
					Utils.mergeCells(spreadsheet.getSelectedSheet(), sel
							.getTop(), sel.getLeft(), sel.getBottom(), sel
							.getRight(), true);
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

	public void onDeleteSheet() {
		final Book book = spreadsheet.getBook();
		if (book == null) {
			return;
		}
		if (book.getNumberOfSheets() == 1) {
			try {
				Messagebox.show("cannot remove last sheet, but you could insert a new sheet, then remove current sheet  ");
			} catch (InterruptedException e) {
			}
			return;
		}

		try {
			Messagebox.show(
				"Do you really want to delete selected sheet \"" + spreadsheet.getSelectedSheet().getSheetName()
				+ "\", those data will be deleted permanently", "",
					Messagebox.YES | Messagebox.NO, "", new EventListener(){

						@Override
						public void onEvent(Event event) throws Exception {
							System.out.println("evt name: " + event.getName());
							if ("onOk" == event.getName()) {
								SheetHelper.deleteSheet(spreadsheet);
							}			
						}});
		} catch (InterruptedException e) {
		}

	}

	
	/**
	 * Switch gridlines setting of spreadsheet
	 * @param event
	 */
	public void onGridlines(ForwardEvent event) {
		Sheet sheet = spreadsheet.getSelectedSheet();
		Utils.getRange(sheet, 0, 0).setDisplayGridlines(!sheet.isDisplayGridlines());
		//TODO avoid use invalidate
		//spreadsheet.invalidate();
	}

	public void openCustomSortDialog() {
		CellHelper.createCustomSortDialog(spreadsheet, Zssapp.getInstance(spreadsheet));
	}

	public void openExportPdfDialog() {
		ExportHelper.doExportToPDF(spreadsheet);
	}

	public void openHyperlinkDialog() {
		Executions.createComponents(Consts._InsertHyperlinkDialog_zul, mainWin, Zssapps.newSpreadsheetArg(spreadsheet));
	}

	public void openInsertFormulaDialog() {
		//Executions.createComponents(Consts._InsertFormulaDialog_zul, mainWin, null);
		Executions.createComponents(Consts._InsertFormulaDialog2_zul, mainWin, null);
	}

	public void openModifyRowHeightDialog() {
		throw new UiException("not implement yet");
	}

	public void openPasteSpecialDialog() {
		Executions.createComponents(Consts._PasteSpecialDialog_zul, mainWin, Zssapps.newSpreadsheetArg(spreadsheet));
	}

	public void toggleFormulaBar() {
		onViewFormulaBar();
	}

	public void openComposeFormulaDialog(FormulaMetaInfo metainfo) {
		HashMap arg = new HashMap();
		arg.put(Consts.KEY_ARG_FORMULA_METAINFO, metainfo);
		Executions.createComponents(Consts._ComposeFormulaDialog_zul, mainWin, arg);
	}

	public void openFormatNumberDialog() {
		Executions.createComponents(Consts._FormatNumberDialog_zul, mainWin, Zssapps.newSpreadsheetArg(spreadsheet));
	}
}