/* MainWindow.java

{{IS_NOTE
	Purpose:
f
	Description:

	History:
		Dec 20, 2007 12:40:46 PM     2007, Created by Dennis.Chenf
		Jan 1, 2009 modified by kinda lu
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.D

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
 */
package org.zkoss.zss.app;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Stack;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.zkoss.image.Image;
import org.zkoss.io.Files;
import org.zkoss.util.media.Media;
import org.zkoss.util.resource.Labels;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Desktop;
import org.zkoss.zk.ui.Execution;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Path;
import org.zkoss.zk.ui.Session;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.event.InputEvent;
import org.zkoss.zk.ui.event.KeyEvent;
import org.zkoss.zk.ui.event.SelectEvent;
import org.zkoss.zk.ui.event.UploadEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.event.CellStyleHelper;
import org.zkoss.zss.app.event.EditHelper;
import org.zkoss.zss.app.event.ExportHelper;
import org.zkoss.zss.app.sort.SortSelector;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.ui.Position;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.CellEvent;
import org.zkoss.zss.ui.event.CellMouseEvent;
import org.zkoss.zss.ui.event.CellSelectionEvent;
import org.zkoss.zss.ui.event.EditboxEditingEvent;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.event.HeaderEvent;
import org.zkoss.zss.ui.event.HeaderMouseEvent;
import org.zkoss.zss.ui.event.StopEditingEvent;
import org.zkoss.zss.ui.impl.MergeMatrixHelper;
import org.zkoss.zss.ui.impl.MergedRect;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zss.ui.sys.SpreadsheetCtrl;
import org.zkoss.zssex.ui.widget.ImageWidget;
import org.zkoss.zul.Borderlayout;
import org.zkoss.zul.Button;
import org.zkoss.zul.Chart;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.Comboitem;
import org.zkoss.zul.Fileupload;
import org.zkoss.zul.Label;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Menu;
import org.zkoss.zul.Menupopup;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.Popup;
import org.zkoss.zul.Radio;
import org.zkoss.zul.Rows;
import org.zkoss.zul.South;
import org.zkoss.zul.Tab;
import org.zkoss.zul.Tabbox;
import org.zkoss.zul.Tabs;
import org.zkoss.zul.Textbox;
import org.zkoss.zul.Toolbarbutton;
import org.zkoss.zul.Window;

/**
 * @author Peter Kuo
 * @modify kinda lu
 */
public class MainWindowCtrl extends GenericForwardComposer {

	private static final long serialVersionUID = 1;
	static int event_x = 200;
	static int event_y = 200;

	private Desktop desktop;
	String hashFilename = null;

	Book book;
	Sheet sheet;

	int lastRow = 0;
	int lastCol = 0;
	// boolean isCut = false;
	boolean isFreezeRow = false;
	boolean isFreezeColumn = false;

	// For fast Icon
	boolean isBold = false;
	boolean isItalic = false;
	boolean isUnderline = false;
	boolean isStrikethrough = false;
	boolean isMergeCell = false;
	boolean isWrapText = false;
	String colorSelectorTarget = "";

	Cell currentEditcell;

	int chartKey = 0;

	ColumnMenuHelper colmh;
	RowMenuHelper rowmh;
	CellMenuHelper cellmh;
	FormatNumberHelper fnh;
	FormulaCategoryHelper fch;
	FileHelper fileh;
	RangeHelper rangeh;
	TabHelper tabh;

	Window mainWin;
	Menu backgroundColorMenu;
	Menu insertImageMenu;
	Menu insertPieChart;

	Textbox formulaEditbox;
	Spreadsheet spreadsheet;
	Tabbox sheetTB;
	Tabs sheetTabs;
	Combobox focusPosition;
	Combobox fontFamilyCombobox;
	Combobox fontSizeCombobox;
	Comboitem lbpos;
	Toolbarbutton boldBtn;
	Toolbarbutton italicBtn;
	Toolbarbutton underlineBtn;
	Toolbarbutton strikethroughBtn;
	Toolbarbutton alignLeftBtn;
	Toolbarbutton alignCenterBtn;
	Toolbarbutton alignRightBtn;
	Colorbutton fontColorBtn;
	Colorbutton backgroundColorBtn;
	Toolbarbutton borderBtn;
	Toolbarbutton mergeCellBtn;
	Toolbarbutton wrapTextBtn;
	Toolbarbutton insertChartBtn;

	South formulaBar;
	Borderlayout topToolbars;

	public static MainWindowCtrl getInstance() {
		Execution exe = Executions.getCurrent();
		return (MainWindowCtrl) exe.getDesktop().getAttribute(
				MainWindowCtrl.class.getCanonicalName());
	}

	public Window getMainWindow() {
		return mainWin;
	}

	public Spreadsheet getSpreadsheet() {
		return spreadsheet;
	}

	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);

		desktop.setAttribute(MainWindowCtrl.class.getCanonicalName(), this);

		ssInit();
		sheetTBInit();

		colmh = new ColumnMenuHelper(spreadsheet);
		rowmh = new RowMenuHelper(spreadsheet);
		cellmh = new CellMenuHelper(spreadsheet);
		fnh = new FormatNumberHelper(spreadsheet);
		fch = new FormulaCategoryHelper(spreadsheet);
		fileh = new FileHelper(spreadsheet);
		rangeh = new RangeHelper(spreadsheet);
		tabh = new TabHelper(spreadsheet);

		formulaEditbox.addEventListener(
				org.zkoss.zk.ui.event.Events.ON_CHANGING, new EventListener() {
					public void onEvent(Event event) throws Exception {
						onEditingFormulaInput(event);
					}
				});

		/**
		 * Retrieve spreadsheet current edit position
		 */
		formulaEditbox.addEventListener(org.zkoss.zk.ui.event.Events.ON_FOCUS,
				new EventListener() {

					public void onEvent(Event event) throws Exception {
						int left = spreadsheet.getSelection().getLeft();
						int top = spreadsheet.getSelection().getTop();
						Sheet sheet = spreadsheet.getSelectedSheet();
						currentEditcell = Utils.getCell(sheet, top, left);
					}
				});

		/**
		 * Reset the spreadsheet edit position
		 */
		formulaEditbox.addEventListener(org.zkoss.zk.ui.event.Events.ON_BLUR,
				new EventListener() {

					public void onEvent(Event event) throws Exception {
						currentEditcell = null;
					}
				});

		formulaEditbox.addEventListener(org.zkoss.zk.ui.event.Events.ON_OK,
				new EventListener() {
					public void onEvent(Event event) throws Exception {

						evalFormula(formulaEditbox.getValue());
						Position pos = spreadsheet.getCellFocus();
						int row = pos.getRow() + 1;
						int col = pos.getColumn();

						spreadsheet.focusTo(row, col);
					}
				});
	}

	private void evalFormula(String input) {
		if (input == null || input.length() == 0)
			return;

		Utils.setEditText(spreadsheet.getSelectedSheet(), spreadsheet
				.getSelection().getTop(), spreadsheet.getSelection().getLeft(),
				input);
	}

	public void ssInit() {
		spreadsheet.setSrcName("Untitled");
		book = spreadsheet.getBook();

		// ADD Event Listener
		spreadsheet.addEventListener(Events.ON_CELL_FOUCSED,
				new EventListener() {
					public void onEvent(Event event) throws Exception {
						FocusedEvent((CellEvent) event);
					}
				});

		spreadsheet.addEventListener(Events.ON_STOP_EDITING,
				new EventListener() {
					public void onEvent(Event event) throws Exception {
						onStopEditingEvent((StopEditingEvent) event);
					}

				});

		spreadsheet.addEventListener(Events.ON_EDITBOX_EDITING,
				new EventListener() {
					public void onEvent(Event event) throws Exception {
						onEditboxEditingEvent((EditboxEditingEvent) event);
					}
				});

		spreadsheet.addEventListener(Events.ON_CELL_RIGHT_CLICK,
				new EventListener() {
					public void onEvent(Event event) throws Exception {
						MouseEvent((CellMouseEvent) event);
					}
				});

		spreadsheet.addEventListener(Events.ON_HEADER_RIGHT_CLICK,
				new EventListener() {
					public void onEvent(Event event) throws Exception {
						onHeaderMouseEvent((HeaderMouseEvent) event);
					}
				});

		spreadsheet.addEventListener(Events.ON_CELL_SELECTION,
				new EventListener() {
					public void onEvent(Event event) throws Exception {
						SelectionEvent((CellSelectionEvent) event);
					}
				});
	}

	public void sheetTBClean() {
		while (sheetTabs.getFirstChild() != null)
			sheetTabs.removeChild(sheetTabs.getFirstChild());
	}

	public void sheetTBInit() {
		int sheetCount;
		sheetCount = book.getNumberOfSheets();
		for (int i = 0; i < sheetCount; i++) {
			sheet = (Sheet) book.getSheetAt(i);
			Tab tab = new Tab(sheet.getSheetName());
			Popup popup = (Popup) mainWin.getFellow("tabPopup");
			tab.setContext(popup);
			tab.setParent(sheetTabs);
		}

		sheetTB.addEventListener(org.zkoss.zk.ui.event.Events.ON_SELECT,
				new EventListener() {

					public void onEvent(Event event) throws Exception {
						onTabboxSelectEvent((SelectEvent) event);
					}
				});
		sheetTB.invalidate();
	}

	// SECTION Sheet Tabbox Event Handler
	void onTabboxSelectEvent(SelectEvent event) {
		spreadsheet.setSelectedSheet(sheetTB.getSelectedTab().getLabel());
	}

	// SECTION Spreadsheet Event Handler
	void MouseEvent(CellMouseEvent event) {
		event_x = event.getClientx();
		event_y = event.getClienty();

		Menupopup cellMenu = (Menupopup) mainWin.getFellow("cellMenu");
		cellMenu.open(event_x + 5, event.getClienty());

		// read cell value to fill fastIcon in contextmenu
		Window win = (Window) mainWin.getFellow("fastIconContextmenu");
		Sheet sheet = event.getSheet();
		Cell cell = Utils.getCell(sheet, lastRow, lastCol);

		// default settings
		Combobox fontFamilyCombobox = (Combobox) win
				.getFellow("_fontFamilyCombobox");
		Combobox fontSizeCombobox = (Combobox) win
				.getFellow("_fontSizeCombobox");
		Toolbarbutton boldBtn = (Toolbarbutton) win.getFellow("_boldBtn");
		Toolbarbutton italicBtn = (Toolbarbutton) win.getFellow("_italicBtn");
		Toolbarbutton alignLeftBtn = (Toolbarbutton) win
				.getFellow("_alignLeftBtn");
		Toolbarbutton alignCenterBtn = (Toolbarbutton) win
				.getFellow("_alignCenterBtn");
		Toolbarbutton alignRightBtn = (Toolbarbutton) win
				.getFellow("_alignRightBtn");
		Colorbutton fontColorBtn = (Colorbutton) win.getFellow("_fontColorBtn");
		Colorbutton backgroundColorBtn = (Colorbutton) win
				.getFellow("_backgroundColorBtn");

		// load default setting
		fontFamilyCombobox.setText("Calibri");
		fontSizeCombobox.setText("12");
		boldBtn.setClass("toolIcon");
		italicBtn.setClass("toolIcon");
		alignLeftBtn.setClass("toolIcon");
		alignCenterBtn.setClass("toolIcon");
		alignRightBtn.setClass("toolIcon");
		fontColorBtn.setColor("#000000");
		backgroundColorBtn.setColor("#FFFFFF");

		if (cell != null) {
			CellStyle cs = cell.getCellStyle();

			if (cs != null) {
				int fontidx = cs.getFontIndex();
				Font font = book.getFontAt((short) fontidx);

				// font family
				fontFamilyCombobox.setText(font.getFontName());

				// font size
				fontSizeCombobox.setText(Integer.toString(font
						.getFontHeightInPoints()));

				// font bold & italic
				isBold = font.getBoldweight() == Font.BOLDWEIGHT_BOLD;
				isItalic = font.getItalic();
				if (isBold)
					boldBtn.setClass("toolIcon clicked");
				if (isItalic)
					italicBtn.setClass("toolIcon clicked");

				// align
				int align = cs.getAlignment();
				switch (align) {
				case CellStyle.ALIGN_LEFT:
					alignLeftBtn.setClass("toolIcon clicked");
					break;
				case CellStyle.ALIGN_CENTER:
				case CellStyle.ALIGN_CENTER_SELECTION:
					alignCenterBtn.setClass("toolIcon clicked");
					break;
				case CellStyle.ALIGN_RIGHT:
					alignRightBtn.setClass("toolIcon clicked");
					break;
				}

				// font color
				String color = BookHelper.getFontHTMLColor(book, font);
				if (color != null && !color.equals(BookHelper.AUTO_COLOR)) {
					fontColorBtn.setColor(color);
				}

				// bg color
				//final int fillColorIdx = cs.getFillForegroundColor();
				//String fcolor = BookHelper.indexToRGB(book, fillColorIdx);
				String fcolor = BookHelper.colorToHTML(book, cs.getFillForegroundColorColor());
				if (fcolor != null && !fcolor.equals(BookHelper.AUTO_COLOR)) {
					backgroundColorBtn.setColor(fcolor);
				}

			}
		}
		win.setPosition("parent");
		win.setLeft(Integer.toString(event_x + 5) + "px");
		win.setTop(Integer.toString(event_y - 100) + "px");
		win.doPopup();
	}

	public void onUpload$insertImageMenu(UploadEvent event) {
		try {
			org.zkoss.util.media.Media media = event.getMedia();
			if (media instanceof org.zkoss.image.Image) {
				ImageWidget image = new ImageWidget();
				image.setContent((Image) media);

				int col = spreadsheet.getSelection().getLeft();
				int row = spreadsheet.getSelection().getTop();
				image.setRow(row);
				image.setColumn(col);

				SpreadsheetCtrl ctrl = (SpreadsheetCtrl) spreadsheet
						.getExtraCtrl();
				ctrl.addWidget(image);
			} else if (media != null) {
				Messagebox.show("Not an image: " + media, "Error",
						Messagebox.OK, Messagebox.ERROR);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	void onHeaderMouseEvent(HeaderMouseEvent event) {
		Menupopup headerMenu;
		if (HeaderEvent.TOP_HEADER == event.getType()) {
			headerMenu = (Menupopup) mainWin.getFellow("columnHeaderMenu");
		} else if ((HeaderEvent.LEFT_HEADER == event.getType())) {
			headerMenu = (Menupopup) mainWin.getFellow("rowHeaderMenu");
		} else {
			return;
		}

		headerMenu.open(event.getClientx(), event.getClienty());
	}

	void FocusedEvent(CellEvent event) {
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
			String editText = Utils.getEditText(cell);
			formulaEditbox.setValue(cell == null ? "" : (editText == null ? ""
					: editText));

			// set to default setting
			fontFamilyCombobox.setText("Calibri");
			fontSizeCombobox.setText("12");
			boldBtn.setClass("toolIcon");
			italicBtn.setClass("toolIcon");
			underlineBtn.setClass("toolIcon");
			strikethroughBtn.setClass("toolIcon");
			alignLeftBtn.setClass("toolIcon");
			alignCenterBtn.setClass("toolIcon");
			alignRightBtn.setClass("toolIcon");
			fontColorBtn.setColor("#000000");
			backgroundColorBtn.setColor("#FFFFFF");
			backgroundColorMenu.setContent("#color=#FFFFFF");
			wrapTextBtn.setClass("toolIcon");

			// read format from cell and assign it to toolbar
			if (cell != null) {
				CellStyle cs = cell.getCellStyle();
				if (cs != null) {
					int fontidx = cs.getFontIndex();
					Font font = book.getFontAt((short) fontidx);
					fontFamilyCombobox.setText(font.getFontName());
					fontSizeCombobox.setText(Integer.toString(font
							.getFontHeightInPoints()));

					// font bold & italic
					isBold = font.getBoldweight() == Font.BOLDWEIGHT_BOLD;
					isItalic = font.getItalic();
					if (isBold)
						boldBtn.setClass("toolIcon clicked");

					if (isItalic)
						italicBtn.setClass("toolIcon clicked");

					// font underline
					if (font.getUnderline() != Font.U_NONE) {
						isUnderline = true;
						underlineBtn.setClass("toolIcon clicked");
					} else {
						isUnderline = false;
					}

					// font stikethrough
					if (font.getStrikeout()) {
						isStrikethrough = true;
						strikethroughBtn.setClass("toolIcon clicked");
					} else {
						isStrikethrough = false;

					}

					// align
					int align = cs.getAlignment();
					switch (align) {
					case CellStyle.ALIGN_LEFT:
						alignLeftBtn.setClass("toolIcon clicked");
						break;
					case CellStyle.ALIGN_CENTER:
					case CellStyle.ALIGN_CENTER_SELECTION:
						alignCenterBtn.setClass("toolIcon clicked");
						break;
					case CellStyle.ALIGN_RIGHT:
						alignRightBtn.setClass("toolIcon clicked");
						break;
					}

					// font color
					String color = BookHelper.getFontHTMLColor(book, font);
					if (color != null && !color.equals(BookHelper.AUTO_COLOR)) {
						fontColorBtn.setColor(color);
					}

					// bg color
//					final int fillColorIdx = cs.getFillForegroundColor();
//					String fcolor = BookHelper.indexToRGB(book, fillColorIdx);
					String fcolor = BookHelper.colorToHTML(book, cs.getFillForegroundColorColor());
					if (fcolor != null && !fcolor.equals(BookHelper.AUTO_COLOR)) {
						backgroundColorBtn.setColor(fcolor);
						backgroundColorMenu.setContent("#color=" + fcolor);
					}

					// merge cell
					// cannot find a way to know the cell is merged or not
					/*
					 * Toolbarbutton mergeCellBtn =
					 * (Toolbarbutton)getFellow("mergeCellBtn");
					 * isMergeCell=format. if(isWrapText){
					 * wrapTextBtn.setClass("toolIcon clicked"); }else{
					 * wrapTextBtn.setClass("toolIcon"); }
					 */

					// wrap text
					isWrapText = cs.getWrapText();
					if (isWrapText) {
						wrapTextBtn.setClass("toolIcon clicked");
					}

					// FontStyle preFontStyle=format.getFontStyle();
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// this is auto pushCellState in spreadsheet itself
	void onStopEditingEvent(StopEditingEvent evt) {
		try {
			// the formula bar input
			formulaEditbox.setValue((String) evt.getEditingValue());

			// to notify all widgets there is a cell changed
			for (int i = 0; i < chartKey; i++) {
				// p("chartKey="+chartKey+" chartWin"+i);
				try {
					Window win = (Window) mainWin.getFellow("chartWin" + i);
					if (win != null) {
						Chart myChart = (Chart) win.getFellow("myChart");
						CellEvent event = new StopEditingEvent(
								org.zkoss.zss.ui.event.Events.ON_STOP_EDITING,
								myChart, evt.getSheet(), evt.getRow(), evt
										.getColumn(), (String) evt.getData());
						org.zkoss.zk.ui.event.Events.postEvent(event);
					}
				} catch (Exception e) {
					// the chart may be deleted
					// e.printStackTrace();
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	void onEditboxEditingEvent(EditboxEditingEvent evt) {
		try {
			formulaEditbox.setValue((String) evt.getEditingValue());
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void SelectionEvent(CellSelectionEvent event) {
		// gonna try to print the current cell is merged or not
		int left = spreadsheet.getSelection().getLeft();
		int right = spreadsheet.getSelection().getRight();
		int top = spreadsheet.getSelection().getTop();
		int bottom = spreadsheet.getSelection().getBottom();

		isMergeCell = false;
		Sheet sheet = spreadsheet.getSelectedSheet();
		MergeMatrixHelper mmhelper = spreadsheet.getMergeMatrixHelper(sheet);
		for (final Iterator iter = mmhelper.getRanges().iterator(); iter
				.hasNext();) {
			MergedRect block = (MergedRect) iter.next();
			int bt = block.getTop();
			int bl = block.getLeft();
			int bb = block.getBottom();
			int br = block.getRight();
			if (left <= bl && top <= bt && right >= br && bottom >= bb) {
				isMergeCell = true;
				break;
			}
		}

		Window win = (Window) mainWin.getFellow("fastIconContextmenu");
		// win.doPopup();
		Toolbarbutton mergeCellBtn2 = (Toolbarbutton) win
				.getFellow("_mergeCellBtn");
		if (isMergeCell) {
			mergeCellBtn.setClass("toolIcon clicked");
			mergeCellBtn2.setClass("toolIcon clicked");
		} else {
			mergeCellBtn.setClass("toolIcon");
			mergeCellBtn2.setClass("toolIcon");
		}
	}

	// SECTION CtrlKeys
	public void onSSCtrlKeys(ForwardEvent event) {
		Event orig = event.getOrigin();
		while (orig instanceof ForwardEvent) {
			orig = ((ForwardEvent) orig).getOrigin();
		}
		if (orig instanceof KeyEvent) {
			_onSSCtrlKeys((KeyEvent) orig);
		} else {
			p("damn not ForwardEvent");
		}
	}

	public void _onSSCtrlKeys(KeyEvent event) {
		char c = (char) event.getKeyCode();
		// p(""+event.getKeyCode()+c+" "+event.isCtrlKey());
		try {
			if (46 == event.getKeyCode()) {// delete
				if (event.isCtrlKey())
					onClearStyle((ForwardEvent) null);
				else
					onClearContent((ForwardEvent) null);
				return;
			}

			if (false == event.isCtrlKey()) {
				p("no ctrl");
				return;
			}

			p("Ctrl+" + c);

			switch (c) {
			case 'X':
				EditHelper.onCut(spreadsheet);
				break;
			case 'C':
				EditHelper.onCopy(spreadsheet);
				break;
			case 'V':
				EditHelper.onPaste(spreadsheet);
				break;
			case 'D':
				onClearContent((ForwardEvent) null);
				break;
			case 'E':// for testing
				spreadsheet.setCellFocus(new Position(2, 2));
				spreadsheet.setSelection(new Rect(2, 2, 4, 4));
				break;
			case 'S':
				fileh.dispatcher("save");
				break;
			// TODO undo
			/*
			 * case 'Z': spreadsheet.undo(); break;
			 */
			// TODO redo
			/*
			 * case 'Y': spreadsheet.redo(); break;
			 */case 'O':
				fileh.selectDispatcher("open");
				break;
			case 'F':

				break;
			case 'B':
				setFontBold();
				break;
			case 'I':
				setFontItalic();
				break;
			case 'U':
				setFontUnderline();
				break;
			default:
				p("ctrl ...?");
				return;
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// SECTION FILE MENU
//	public void onFileSelect(ForwardEvent event) {
//		fileh.selectDispatcher((String) event.getData());
//	}
	public void onExport(ForwardEvent event) {
		ExportHelper.onExport(event);
	}

	public void onCellMenu(ForwardEvent event) {
		cellmh.dispatcher((String) event.getData());
	}

	public void onFileOK(ForwardEvent event) {
		fileh.dispatcher((String) event.getData());
	}

	public void onFileNew(ForwardEvent event) {
		openNewFileInFS();
	}

	public void onFileSaveClose(ForwardEvent event) {
		fileh.selectDispatcher("save");
		openNewFileInFS();
	}

	public void onOpenFile(ForwardEvent event) {
		Listbox flo_files = (Listbox) Path
				.getComponent("//p1/mainWin/fileOpenWin/flo_files");
		String filename = null;
		if (flo_files.getSelectedItem() != null) {
			filename = flo_files.getSelectedItem().getLabel();
			p(filename);
			openFileInFS(filename);
		}

		Window fileOpenWin = (Window) Path
				.getComponent("//p1/mainWin/fileOpenWin");
		fileOpenWin.setVisible(false);
		// spreadsheet.conCurrentEditing(this);
		spreadsheet.invalidate();
	}

	public void onDeleteFile(ForwardEvent event) {
		Listbox fld_files = (Listbox) Path
				.getComponent("//p1/mainWin/fileDeleteWin/fld_files");

		String filename = fld_files.getSelectedItem().getLabel();
		fileh.deleteFile(filename);

		Window fileDeleteWin = (Window) Path
				.getComponent("//p1/mainWin/fileDeleteWin");
		fileDeleteWin.setVisible(false);

		if (filename.equals(spreadsheet.getSrc()))
			onFileNew(null);
	}

	public void onExportFile(ForwardEvent event) {
		Listbox fle_files = (Listbox) Path
				.getComponent("//p1/mainWin/fileExportWin/fle_files");

		String filename = fle_files.getSelectedItem().getLabel();
		exportFile(filename);

		Window fileExportWin = (Window) Path
				.getComponent("//p1/mainWin/fileExportWin");
		fileExportWin.setVisible(false);
	}

	public void openFileInFS(String filename) {
		HashMap hm = fileh.readMetafile();
		File file;
		if (filename != null && hm.get(filename) != null) {
			hashFilename = (String) ((Object[]) hm.get(filename))[1];
		} else {
			hashFilename = filename;
			filename = spreadsheet.getSrc();
		}
		p(fileh.xlsDir + hashFilename);
		file = new File(fileh.xlsDir + hashFilename);
		Label label = (Label) Path.getComponent("//p1/mainWin/openingFilename");
		label.setValue(filename);
		if (file == null)
			System.out.println("cannot found file: " + file.getName());
		try {
			openSpreadsheetFromStream(new FileInputStream(file), filename);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public void onImportAndOpenFile(ForwardEvent event) {
		openFileInFS(onImportFile(null));
	}

	public String onImportFile(ForwardEvent event) {
		String metaStr = null;
		String filename = null;
		try {
			Object ss = Fileupload.get();
			if (ss != null) {
				InputStream iStream = ((Media) ss).getStreamData();
				filename = ((Media) ss).getName();
				long millis = System.currentTimeMillis();
				String hashFilename = millis + ".xls";

				Files.copy(new File(fileh.xlsDir + hashFilename), iStream);

				FileOutputStream metaStream = new FileOutputStream(fileh.xlsDir
						+ "metaFile", true);

				metaStr = new String(millis + "\n" + filename + "\n"
						+ hashFilename + "\n");
				// System.out.println(fileh.xlsDir);
				metaStream.write(metaStr.getBytes());
				iStream.close();
				metaStream.close();
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		return filename;
	}

	public void openNewFileInFS() {

		File file = new File(fileh.xlsDir + "Untitled");

		try {
			openSpreadsheetFromStream(new FileInputStream(file), "Untitled");
			Label label = (Label) Path
					.getComponent("//p1/mainWin/openingFilename");
			label.setValue("Untitled");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

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

		String printKey = "" + System.currentTimeMillis();
		Session session = Executions.getCurrent().getDesktop().getSession();
		session.setAttribute("zssFromHi" + printKey, spreadsheet);

		Window win = (Window) mainWin.getFellow("menuPrintWin");
		Button printBtn = (Button) win.getFellow("printBtn");
		printBtn.setHref("print.zul?printKey=" + printKey);
		printBtn.setTooltip("print.zul?printKey=" + printKey);

		win.setPosition("parent");
		win.setTop("100px");
		win.setLeft("100px");
		win.doPopup();
	}

	public void onRevision(ForwardEvent event) {
		reloadRevisionMenu();
		Window win = (Window) mainWin.getFellow("revisionWin");
		win.setPosition("parent");
		win.setLeft("250px");
		win.setTop("250px");
		win.doPopup();
	}

	// SECTION EDIT MENU
	public void onEditUndo(ForwardEvent event) {
		// TODO undo
		throw new UiException("Undo not implmented yet");
		// spreadsheet.undo();
	}

	public void onEditRedo(ForwardEvent event) {
		// TODO redo
		throw new UiException("Redo not implemented yet");
		// spreadsheet.redo();
	}

	public void onEditCut(ForwardEvent event) {
		EditHelper.onCut(spreadsheet);
	}

	public void onEditCopy(ForwardEvent event) {
		EditHelper.onCopy(spreadsheet);
	}

	public void onEditPaste(ForwardEvent event) {
		EditHelper.onPaste(spreadsheet);
	}

	public void onClearContent(ForwardEvent event) {
		// TODO undo/redo
		// spreadsheet.pushCellState();
		onClearContent();
	}

	public void onClearContent() {
		try {
			int left = spreadsheet.getSelection().getLeft();
			int right = spreadsheet.getSelection().getRight();
			int top = spreadsheet.getSelection().getTop();
			int bottom = spreadsheet.getSelection().getBottom();
			Sheet sheet = spreadsheet.getSelectedSheet();
			Cell cell = null;
			for (int row = top; row <= bottom; row++)
				for (int col = left; col <= right; col++) {
					cell = Utils.getCell(sheet, row, col);
					if (cell != null) {
						Utils.setEditText(cell, null);
					}
				}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void onClearStyle(ForwardEvent event) {
		// TODO undo/redo
		// spreadsheet.pushCellState();
		onClearStyle();
	}

	public void onClearStyle() {
		int left = spreadsheet.getSelection().getLeft();
		int right = spreadsheet.getSelection().getRight();
		int top = spreadsheet.getSelection().getTop();
		int bottom = spreadsheet.getSelection().getBottom();
		Sheet sheet = spreadsheet.getSelectedSheet();
		Cell cell = null;
		for (int row = top; row <= bottom; row++) {
			for (int col = left; col <= right; col++) {
				cell = Utils.getCell(sheet, row, col);
				if (cell != null) {
					CellStyle newCellStyle = book.createCellStyle();
					Range rng = Utils.getRange(sheet, row, col);
					rng.setStyle(newCellStyle);
				}
			}
		}
	}

	public void onClearBoth(ForwardEvent event) {
		// TODO undo/redo
		// spreadsheet.pushCellState();
		onClearStyle();
		onClearContent();
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
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
			// TODO undo/redo
			// spreadsheet.pushDeleteRowColState(-1, top, -1, bottom);
			Utils.deleteRows(sheet, top, bottom);
			spreadsheet.setSelection(rect);
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
			spreadsheet.setSelection(rect);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	//TODO: Do not use hard code
	public void onViewFormulaBar(ForwardEvent event) {
		if (formulaBar.getHeight() != "0px") {
			formulaBar.setSize("0px");
			topToolbars.setHeight("58px");
		} else {
			topToolbars.setHeight("85px");
			formulaBar.setSize("23px");
		}
	}

	public void onViewFreezeRows(ForwardEvent event) {
		spreadsheet.setRowfreeze(Integer.parseInt((String) event.getData()) - 1);
	}

	public void onViewFreezeCols(ForwardEvent event) {
		spreadsheet.setColumnfreeze(Integer.parseInt((String) event.getData()) - 1);
	}

	public void onViewUnfreezeRows(ForwardEvent event) {
		spreadsheet.setRowfreeze(-1);
	}

	public void onViewUnfreezeCols(ForwardEvent event) {
		spreadsheet.setColumnfreeze(-1);
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
			rect.setTop(bottom + 1);
			rect.setBottom(bottom + bottom - top + 1);
			spreadsheet
					.setCellFocus(new Position(rect.getTop(), rect.getLeft()));
			spreadsheet.setSelection(rect);

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
			rect.setLeft(right + 1);
			rect.setRight(right + right - left + 1);
			spreadsheet.setCellFocus(new Position(rect.getTop(), rect.getLeft()));
			spreadsheet.setSelection(rect);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void onInsertSheet() {
		onInsertSheet(true);
	}

	public void onInsertSheet(boolean isNotifyAll) {
		int sheetCount = book.getNumberOfSheets();
		Sheet addedSheet = book.createSheet("sheet " + (sheetCount + 1));

		Tab tab = new Tab(addedSheet.getSheetName());
		Popup popup = (Popup) Path.getComponent("//p1/mainWin/tabPopup");
		tab.setContext(popup);
		tab.setParent(sheetTB.getFellow("sheetTabs"));
	}

	// SECTION HELP MENU
	public void onHelpCheatsheet(ForwardEvent event) {
		Window win = (Window) mainWin.getFellow("cheatsheet");
		win.setPosition("parent");
		win.setLeft("327px");
		win.setTop("124px");
		win.doPopup();
	}

	// SECTION Formula bar
	public void onFormulaBarBtn(ForwardEvent event) {
		event_x = 207;
		event_y = 101;
		onFormulaPopup();
	}

	public void onFormulaListOK() {
		Combobox formulaList = (Combobox) Path
				.getComponent("//p1/mainWin/formulaList");

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

	public void onFormulaCategoryOK() {
		Window win = (Window) mainWin.getFellow("formulaCategory");
		win.setVisible(false);
		fch.onOK();
	}

	public void onEditingFormulaInput(Event event) {
		if (currentEditcell == null)
			return;
		if (currentEditcell.getCellType() == Cell.CELL_TYPE_FORMULA
				&& currentEditcell.getCellFormula() != null)
			return;

		Utils.setEditText(currentEditcell, ((InputEvent) event).getValue());
	}

	/**
	 * Returns the font size in point from the font size combobox
	 * @return short, font size height
	 */
	private short getFontHeightInPoints() {
		String fontSizeStr = this.fontSizeCombobox.getSelectedItem().getLabel();
		return (short) (Integer.parseInt(fontSizeStr) * 20);
	}

	/**
	 * Set font style
	 * 
	 * @param sheet
	 * @param rect
	 */
	private void changeFont(Sheet sheet, Rect rect) {
		// font
		String fontName = (String) this.fontFamilyCombobox.getSelectedItem().getLabel();
		short boldWeight = isBold ? 
				Font.BOLDWEIGHT_BOLD : Font.BOLDWEIGHT_NORMAL;
		byte underline = isUnderline ? Font.U_SINGLE : Font.U_NONE;
		Color color = BookHelper.HTMLToColor(book, fontColorBtn.getColor());
		Utils.setFont(sheet, rect, boldWeight, color, getFontHeightInPoints(),
				fontName, isItalic, isStrikethrough, Font.SS_NONE, underline);
	}

	public void setFontFamily(String fontName) {
		List items = fontFamilyCombobox.getItems();
		boolean findFont = false;
		for (int i = 0; i < items.size(); i++) {
			Comboitem item = (Comboitem) items.get(i);
			if (item.getLabel().equals(fontName)) {
				fontFamilyCombobox.setSelectedItem(item);
				findFont = true;
				break;
			}
		}
		if (findFont)
			Utils.setFontFamily(spreadsheet.getSelectedSheet(), getSpreadsheetMaxSelection(), fontName);
	}

	// SECTION FastIcon Toolbar
	public void onFontFamilySelect(ForwardEvent event) {
		// TODO undo/redo
		// spreadsheet.pushCellState();
		Combobox fontFamily = null;
		String font;
		Window win = (Window) mainWin.getFellow("fastIconContextmenu");
		if (win.isVisible()) {
			win.setVisible(false);
			fontFamily = (Combobox) win.getFellow("_fontFamilyCombobox");
		} else
			fontFamily = this.fontFamilyCombobox;
		if (event.getData() == null)
			font = fontFamily.getSelectedItem().getLabel();
		else
			font = (String) event.getData();
		this.fontFamilyCombobox.setSelectedIndex((fontFamily.getSelectedIndex()));
		Utils.setFontFamily(spreadsheet.getSelectedSheet(), getSpreadsheetMaxSelection(), font);
	}

	public void setFontSize(String fontSize) {
		List items = fontSizeCombobox.getItems();
		boolean findSize = false;
		for (int i = 0; i < items.size(); i++) {
			Comboitem item = (Comboitem) items.get(i);
			if (item.getLabel().equals(fontSize)) {
				fontSizeCombobox.setSelectedItem(item);
				findSize = true;
				break;
			}
		}
		if (findSize)
			Utils.setFontHeight(spreadsheet.getSelectedSheet(), getSpreadsheetMaxSelection(), getFontHeightInPoints());
	}

	public void onFontSizeSelect(ForwardEvent event) {
		// TODO undo/redo
		// spreadsheet.pushCellState();
		Combobox fontSize;
		String size;

		Window win = (Window) mainWin.getFellow("fastIconContextmenu");
		if (win.isVisible()) {
			win.setVisible(false);
			fontSize = (Combobox) win.getFellow("_fontSizeCombobox");
		} else {
			fontSize = fontSizeCombobox;
		}
		if (event.getData() == null)
			size = fontSize.getSelectedItem().getLabel();
		else
			size = (String) event.getData();

		fontSizeCombobox.setSelectedIndex((fontSize.getSelectedIndex()));
		Utils.setFontHeight(spreadsheet.getSelectedSheet(), getSpreadsheetMaxSelection(), getFontHeightInPoints());
	}

	public void onFontBoldClick(ForwardEvent event) {
		setFontBold();
	}

	public void setFontBold() {
		// TODO undo/redo
		// spreadsheet.pushCellState();

		//switch font bold and set sclass
		boldBtn.setClass((isBold = !isBold) ? "toolIcon clicked" : "toolIcon");
		Utils.setFontBold(spreadsheet.getSelectedSheet(), getSpreadsheetMaxSelection(), isBold);
	}

	public void onFontItalicClick(ForwardEvent event) {
		setFontItalic();
	}

	public void setFontItalic() {
		// TODO undo/redo
		// spreadsheet.pushCellState();
		
		//switch italic and set sclass
		italicBtn.setClass((isItalic = !isItalic) ? "toolIcon clicked" : "toolIcon");
		Utils.setFontItalic(spreadsheet.getSelectedSheet(), getSpreadsheetMaxSelection(), isItalic);
	}

	public void onFontUnderlineClick(ForwardEvent event) {
		setFontUnderline();
	}
	
	
	public void setFontUnderline() {
		// TODO undo/redo
		// spreadsheet.pushCellState();
		
		//switch underline and set sclass
		underlineBtn.setClass((isUnderline = !isUnderline) ? "toolIcon clicked" : "toolIcon");
		Utils.setFontUnderline(spreadsheet.getSelectedSheet(), getSpreadsheetMaxSelection(), isUnderline ? Font.U_SINGLE : Font.U_NONE);
	}

	public void onFontStrikethroughClick(ForwardEvent event) {
		setFontStrikethrough();
	}
	
	public void setFontStrikethrough() {
		// TODO undo/redo
		// spreadsheet.pushCellState();
		
		//switch strike and set sclass
		strikethroughBtn.setClass((isStrikethrough = !isStrikethrough) ? "toolIcon clicked" : "toolIcon");
		Utils.setFontStrikeout(spreadsheet.getSelectedSheet(), getSpreadsheetMaxSelection(), isStrikethrough);
	}

	public void onAlignHorizontalClick(ForwardEvent event) {
		// TODO undo/redo
		// spreadsheet.pushCellState();

		mainWin.getFellow("fastIconContextmenu").setVisible(false);
		String alignStr = (String) event.getData();

		short align = CellStyle.ALIGN_GENERAL;

		if (alignStr.equals("left")) {
			alignLeftBtn.setClass("toolIcon clicked");
			align = CellStyle.ALIGN_LEFT;
		} else
			alignLeftBtn.setClass("toolIcon");

		if (alignStr.equals("center")) {
			alignCenterBtn.setClass("toolIcon clicked");
			align = CellStyle.ALIGN_CENTER;
		} else
			alignCenterBtn.setClass("toolIcon");

		if (alignStr.equals("right")) {
			alignRightBtn.setClass("toolIcon clicked");
			align = CellStyle.ALIGN_RIGHT;
		} else
			alignRightBtn.setClass("toolIcon");

		Utils.setAlignment(spreadsheet.getSelectedSheet(), getSpreadsheetMaxSelection(), align);
	}

	public void onChange$fontColorBtn() {
		this.setFontColor(fontColorBtn.getColor());
	}

	public void onChange$backgroundColorBtn() {
		this.setBackgroundColor(backgroundColorBtn.getColor());
	}

	public void onChange$backgroundColorMenu(InputEvent event) {
		setBackgroundColor(event.getValue());
	}

	public void setFontColor(String color) {
		if (color == null)
			return;

		fontColorBtn.setColor(color);
		changeFont(spreadsheet.getSelectedSheet(), getSpreadsheetMaxSelection());
	}

	public void setBackgroundColor(String color) {
		if (color == null)
			return;

		backgroundColorBtn.setColor(color);
		Utils.setBackgroundColor(spreadsheet.getSelectedSheet(), getSpreadsheetMaxSelection(), backgroundColorBtn.getColor());
	}

	/**
	 * 
	 * @param event
	 */
	public void onWrapTextClick(ForwardEvent event) {
		// TODO remove me, wrap text
		throw new UiException("wrap text is implmented yet");
		/*
		 * spreadsheet.pushCellState(); try{ isWrapText=!isWrapText;
		 * if(isWrapText){ wrapTextBtn.setClass("toolIcon clicked"); }else{
		 * wrapTextBtn.setClass("toolIcon"); }
		 * 
		 * int left = spreadsheet.getSelection().getLeft(); int right =
		 * spreadsheet.getSelection().getRight(); int top =
		 * spreadsheet.getSelection().getTop(); int bottom =
		 * spreadsheet.getSelection().getBottom(); Sheet sheet =
		 * spreadsheet.getSelectedSheet(); for (int row = top; row <= bottom;
		 * row++) for (int col = left; col <= right; col++){
		 * Styles.setTextWrap(sheet, row, col, isWrapText); } }catch(Exception
		 * e){ e.printStackTrace(); }
		 */
	}

	public void onMergeCellClick(ForwardEvent event) {
		try {
			mainWin.getFellow("fastIconContextmenu").setVisible(false);

			// isMergeCell should read from the cell
			// and search all over the cell if any one is merged then unmerge
			// all
			isMergeCell = !isMergeCell;
			if (isMergeCell) {
				mergeCellBtn.setClass("toolIcon clicked");
			} else {
				mergeCellBtn.setClass("toolIcon");
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

	private void sortAscending() {
		Utils.sort(spreadsheet.getSelectedSheet(), getSpreadsheetMaxSelection(),
				null, null, null, false, false, false);
	}

	private void sortDescending() {
		Utils.sort(spreadsheet.getSelectedSheet(), getSpreadsheetMaxSelection(),
				null, new boolean[] { true }, null, false, false, false);
	}

	/**
	 * Execute sort function base on event's parameter
	 * <p>
	 * Parameter can indicate on either ascending sort or descending sort or
	 * custom sort
	 * <p>
	 * If parameter indicate custom sort, will open a custom sort window dialog
	 * 
	 * @param event
	 */
	public void onSortSelector(ForwardEvent event) {
		String param = (String) event.getData();
		if (param == null)
			return;
		if (param.equals(Labels.getLabel("sort.custom"))) {
			HashMap arg = new HashMap();
			arg.put("spreadsheet", spreadsheet);
			Executions.createComponents("/menus/sort/customSort.zul", mainWin,
					arg);
		} else {
			if (SortSelector.getSortOrder(param))
				sortAscending();
			else
				sortDescending();
		}
	}

	public void onPasteSelector(ForwardEvent event) {
		EditHelper.onPasteEventHandler(event);
	}

	public void onBorderClick(ForwardEvent event) {
		try {
			Window win = (Window) mainWin.getFellow("borderSelector");
			win.setPosition("parent");
			if (event.getData() != null
					&& ((String) event.getData()).equals("toolbar")) {
				win.setLeft("668px");
				win.setTop("52px");
			} else {
				mainWin.getFellow("fastIconContextmenu").setVisible(false);
				((Window) mainWin.getFellow("fastIconContextmenu")).doPopup();
				win.setLeft(80 + event_x + "px");
				win.setTop(-108 + event_y + "px");
			}
			win.doPopup();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void onBorderSelector(ForwardEvent event) {
		// TODO: undo/redo
		CellStyleHelper.onBorderEventHandler(event);
	}

	public void onFormatOK(ForwardEvent event) {
		// p("formatOk"+(String)event.getData());
		fnh.onOK();
	}

	public void onColumnHeaderMenu(ForwardEvent event) {
		colmh.dispatcher((String) event.getData());
	}

	public void onRowHeaderMenu(ForwardEvent event) {
		rowmh.dispatcher((String) event.getData());
	}

	public void Menu(ForwardEvent event) {
		cellmh.dispatcher((String) event.getData());
	}

	public void onRange(ForwardEvent event) {
		rangeh.dispatcher((String) event.getData());
	}

	public void onTab(ForwardEvent event) {
		tabh.dispatcher((String) event.getData());
	}

	// SECTION Utility Functions
	public void p(String s) {
		// System.out.println(spreadsheet.getEditorName()+": "+s);
	}

	public void onFormulaPopup() {
		Window win = (Window) mainWin.getFellow("formulaCategory");
		win.doHighlighted();
	}

	public void onFormatPopup() {
		Window win = (Window) mainWin.getFellow("formatNumberWin");
		try {
			// the set position only work at the second time?
			win.setPosition("parent");
			win.setLeft(event_x + "px");
			win.setTop(event_y + "px");
			win.doPopup();// Modal();
		} catch (Exception e) {
			e.printStackTrace();
		}
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

		try {
			// read the metafile
			BufferedReader bReader;
			bReader = new BufferedReader(new FileReader(fileh.xlsDir
					+ "metaFile"));
			String line = null;
			Stack stack = new Stack();

			String timeStr, filename, hashFilename;
			while (true) {
				timeStr = bReader.readLine();
				if (timeStr == null) {
					break;
				}
				filename = bReader.readLine();
				if (filename == null) {
					System.out
							.println("Warning: filename cannot read from metaFile");
					break;
				}
				hashFilename = bReader.readLine();
				if (hashFilename == null) {
					System.out
							.println("Warning: hashFilename cannot read from metaFile");
					break;
				}
				if (spreadsheet.getSrc().equals(filename)) {
					stack.add(timeStr);
					stack.add(filename);
					stack.add(hashFilename);
				}
			}
			bReader.close();

			// put read data to Menu
			org.zkoss.zul.Row newRow;
			Rows revisionRows = (Rows) mainWin.getFellow("revisionWin")
					.getFellow("revisionRows");
			List rowsChildList = revisionRows.getChildren();
			while (!rowsChildList.isEmpty())
				revisionRows.removeChild((Component) rowsChildList.get(0));
			while (!stack.isEmpty()) {

				hashFilename = (String) stack.pop();
				filename = (String) stack.pop();
				timeStr = (String) stack.pop();
				String dateStr = new Date(Long.parseLong(timeStr)).toString();

				newRow = new org.zkoss.zul.Row();
				Radio radio = new Radio();
				radio.setAttribute("value", hashFilename);

				if (hashFilename.equals(this.hashFilename)) {
					newRow.appendChild(new Label("current"));
					newRow.setStyle("background:rgb(250,230,180) none");
				} else {
					newRow.appendChild(radio);
					// p("set hashFilename: "+hashFilename);
				}
				newRow.appendChild(new Label(dateStr));
				newRow.appendChild(new Label("test name"));
				newRow.appendChild(new Label("no comment"));

				revisionRows.appendChild(newRow);
				// p(""+radio.getAttribute("value"));
			}
			// close the jdbc connection

		} catch (Exception ex) {
			p("exception: " + ex.getMessage());

		}
	}

	public void onRevisionOK(ForwardEvent event) {
		String filename = null;

		Rows revisionRows = (Rows) mainWin.getFellow("revisionWin").getFellow(
				"revisionRows");
		List rowList = revisionRows.getChildren();
		for (int i = 0; i < rowList.size(); i++) {
			Component tmpComponent = ((Component) rowList.get(i))
					.getFirstChild();
			if (tmpComponent instanceof Radio
					&& ((Radio) tmpComponent).isChecked()) {
				filename = (String) ((Radio) tmpComponent)
						.getAttribute("value");
				// p("onRevision: "+filename);
			}
		}

		openFileInFS(filename);

		Window revisionWin = (Window) Path
				.getComponent("//p1/mainWin/revisionWin");
		revisionWin.setVisible(false);
		spreadsheet.invalidate();
		// spreadsheet.notifyRevision();
	}

	public void openSpreadsheetFromStream(InputStream iStream, String src) {

		spreadsheet.setBookFromStream(iStream, src);

		book = spreadsheet.getBook();
		sheetTBClean();
		sheetTBInit();
		sheetTB.addEventListener(org.zkoss.zk.ui.event.Events.ON_SELECT,
				new EventListener() {
					public void onEvent(Event event) throws Exception {
						onTabboxSelectEvent((SelectEvent) event);
					}
				});

		spreadsheet.setSelectedSheet(((Tab) sheetTB.getTabs().getFirstChild())
				.getLabel());
		spreadsheet.invalidate();
	}

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
		// TODO Check if there is the last sheet
		if (spreadsheet.getBook().getNumberOfSheets() == 1) {
			try {
				Messagebox
						.show("cannot remove last sheet, but you could insert a new sheet, then remove current sheet  ");
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			return;
		}

		// TODO show the hint "Are you sure?"
		String name = spreadsheet.getSelectedSheet().getSheetName();
		int buttons = Messagebox.YES + Messagebox.NO;
		int result = -1;
		try {
			result = Messagebox.show(
					"Do you really want to delete selected sheet \"" + name
							+ "\", those data will be deleted permanently", "",
					buttons, "");
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		if (result == Messagebox.YES) {
			int index = book.getSheetIndex(spreadsheet.getSelectedSheet());
			onDeleteSheet(index);
		}

	}

	public void onDeleteSheet(int index) {
		// clean the related state in the undo/redo stack
		// TODO undo/redo
		// spreadsheet.cleanRelatedState(spreadsheet.getSheet(index));

		Book book = spreadsheet.getBook();
		book.removeSheetAt(index);

		int sheetCount = book.getNumberOfSheets();
		Tabbox sheetTB = (Tabbox) Path.getComponent("//p1/mainWin/sheetTB");

		redrawTab();

		if (sheetCount > 0) {
			sheetTB.setSelectedIndex(0);
			spreadsheet.setSelectedSheet(sheetTB.getSelectedTab().getLabel());
		} else {
			spreadsheet.invalidate();
			// TODO: after remove all sheets, why is there still sheet1's
			// content?
		}
	}

	public void redrawTab() {
		Sheet sheet;
		Tabbox sheetTB = (Tabbox) Path.getComponent("//p1/mainWin/sheetTB");
		sheetTB.getFellow("sheetTabs").getChildren().clear();
		Book book = spreadsheet.getBook();
		int sheetCount = book.getNumberOfSheets();
		for (int i = 0; i < sheetCount; i++) {
			sheet = (Sheet) book.getSheetAt(i);
			Tab tab = new Tab(sheet.getSheetName());
			Popup popup = (Popup) Path.getComponent("//p1/mainWin/tabPopup");
			tab.setContext(popup);
			tab.setParent(sheetTB.getFellow("sheetTabs"));
		}
	}
	
	private Rect getSpreadsheetMaxSelection() {
		Rect selection = spreadsheet.getSelection();
		if (selection.getBottom() >= spreadsheet.getMaxrows())
			selection.setBottom(spreadsheet.getMaxrows() - 1);
		if (selection.getRight() >= spreadsheet.getMaxcolumns())
			selection.setRight(spreadsheet.getMaxcolumns() - 1);
		return selection;
	}
}