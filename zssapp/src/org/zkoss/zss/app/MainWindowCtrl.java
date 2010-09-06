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
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.zkoss.image.Image;
import org.zkoss.io.Files;
import org.zkoss.util.media.Media;
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
import org.zkoss.zul.Menuitem;
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
	Sheet remSheet;

	int lastRow = 0;
	int lastCol = 0;
	boolean isCut = false;
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

	int chartKey=0;

	// Sheet remSheet;
	// String focusSheetName;

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
	Menuitem sortAscendingMenu;
	Menuitem sortDescendingMenu;
	Menuitem customSort;
	Menuitem pasteSpecial;
	
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
	Toolbarbutton sortAscending;
	Toolbarbutton sortDescending;
	Toolbarbutton copySelection;
	Toolbarbutton pasteSelection;

	South formulaBar;
	Borderlayout topToolbars; 
	
	public static MainWindowCtrl getInstance() {
		Execution exe = Executions.getCurrent();
		return (MainWindowCtrl)exe.getDesktop().getAttribute(MainWindowCtrl.class.getCanonicalName());
	}

	public Window getMainWindow() {
		return mainWin;
	}
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		desktop.setAttribute(MainWindowCtrl.class.getCanonicalName() , this);
		
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
		
		formulaEditbox.addEventListener(org.zkoss.zk.ui.event.Events.ON_CHANGING, new EventListener(){
			public void onEvent(Event event) throws Exception {
				onEditingFormulaInput(event);
			}
		});
		
		/**
		 * Retrieve spreadsheet current edit position
		 */
		formulaEditbox.addEventListener(org.zkoss.zk.ui.event.Events.ON_FOCUS, new EventListener(){

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
		formulaEditbox.addEventListener(org.zkoss.zk.ui.event.Events.ON_BLUR, new EventListener(){

			public void onEvent(Event event) throws Exception {
				currentEditcell = null;
			}			
		});
		
		formulaEditbox.addEventListener(org.zkoss.zk.ui.event.Events.ON_OK, new EventListener(){
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
		
		int left = spreadsheet.getSelection().getLeft();
		int top = spreadsheet.getSelection().getTop();
		Sheet sheet = spreadsheet.getSelectedSheet();
		
		Utils.setEditText(sheet, top, left, input);
		System.out.println("*** setEditText ***:"+input);
/*		Cell cell = Utils.getCell(sheet, top, left);
		
		String f = cell.getCellFormula();
		if (input.startsWith("=") || f != null) {
			cell.setCellFormula(input.substring(1));
			System.out.println("*** setFormula***:"+input);
		}
*/	}

	public void ssInit() {
		spreadsheet.setSrcName("Untitled");
		book = spreadsheet.getBook();
		
		
		//ADD Event Listener
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
	

	
	public void sheetTBClean(){
		while (sheetTabs.getFirstChild() != null)
			sheetTabs.removeChild(sheetTabs.getFirstChild());
	}
	
	public void sheetTBInit(){
		int sheetCount;
		sheetCount= book.getNumberOfSheets();
		for (int i = 0; i < sheetCount; i++) {
			sheet = (Sheet) book.getSheetAt(i);
			Tab tab = new Tab(sheet.getSheetName());
			Popup popup = (Popup) mainWin.getFellow("tabPopup");
			tab.setContext(popup);
			tab.setParent(sheetTabs);
		}

		sheetTB.addEventListener(org.zkoss.zk.ui.event.Events.ON_SELECT, new EventListener() {
			
			public void onEvent(Event event) throws Exception {
				onTabboxSelectEvent((SelectEvent) event);
			}
		});
		sheetTB.invalidate();
	}
	

	//SECTION Sheet Tabbox Event Handler
	void onTabboxSelectEvent(SelectEvent event) {
		spreadsheet.setSelectedSheet(sheetTB.getSelectedTab().getLabel());
	}
	
	
	
	//SECTION Spreadsheet Event Handler
	void MouseEvent(CellMouseEvent event) {
		event_x = event.getClientx();
		event_y = event.getClienty();
		Menupopup cellMenu = (Menupopup) mainWin.getFellow("cellMenu");
		cellMenu.open(event_x + 5, event.getClienty());

		//read cell value to fill fastIcon in contextmenu
		Window win = (Window) mainWin.getFellow("fastIconContextmenu");
		Sheet sheet = event.getSheet();
		Cell cell = Utils.getCell(sheet, lastRow, lastCol);
		
		//default settings
		Combobox fontFamilyCombobox = (Combobox)win.getFellow("_fontFamilyCombobox");
		Combobox fontSizeCombobox = (Combobox)win.getFellow("_fontSizeCombobox");
		Toolbarbutton boldBtn = (Toolbarbutton)win.getFellow("_boldBtn");
		Toolbarbutton italicBtn = (Toolbarbutton)win.getFellow("_italicBtn");	
		Toolbarbutton alignLeftBtn=(Toolbarbutton)win.getFellow("_alignLeftBtn");
		Toolbarbutton alignCenterBtn=(Toolbarbutton)win.getFellow("_alignCenterBtn");
		Toolbarbutton alignRightBtn=(Toolbarbutton)win.getFellow("_alignRightBtn");
		Colorbutton fontColorBtn = (Colorbutton)win.getFellow("_fontColorBtn");
		Colorbutton backgroundColorBtn = (Colorbutton)win.getFellow("_backgroundColorBtn");
		
		//load default setting
		fontFamilyCombobox.setText("Calibri");
		fontSizeCombobox.setText("12");
		boldBtn.setClass("toolIcon");
		italicBtn.setClass("toolIcon");
		alignLeftBtn.setClass("toolIcon");
		alignCenterBtn.setClass("toolIcon");
		alignRightBtn.setClass("toolIcon");
		fontColorBtn.setColor("#000000");
		backgroundColorBtn.setColor("#FFFFFF");
		
		//p("cell==null");
		if(cell!=null){
			//p("cell!=null");
			CellStyle cs = cell.getCellStyle();
			
			if (cs != null) {
				int fontidx  = cs.getFontIndex();
				Font font = book.getFontAt((short)fontidx);
				
				//font family
				fontFamilyCombobox.setText(font.getFontName());
				
				//font size
				fontSizeCombobox.setText(Integer.toString(font.getFontHeightInPoints()));
				
				//font bold & italic
				isBold = font.getBoldweight() == Font.BOLDWEIGHT_BOLD;
				isItalic = font.getItalic();
				if(isBold)
					boldBtn.setClass("toolIcon clicked");
				if(isItalic)
					italicBtn.setClass("toolIcon clicked");
		
				//align
				int align  = cs.getAlignment();
				switch(align) {
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
				
				//font color
				String color = BookHelper.getFontColor(book, font);
				if (color != null && !color.equals(BookHelper.AUTO_COLOR)) {
					fontColorBtn.setColor(color);
				}
				
				// bg color
				final int fillColorIdx = cs.getFillForegroundColor();
				String fcolor = BookHelper.indexToRGB(book, fillColorIdx);
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

				SpreadsheetCtrl ctrl = (SpreadsheetCtrl) spreadsheet.getExtraCtrl();
				ctrl.addWidget(image);
			} else if (media != null) {
				Messagebox.show("Not an image: " + media, "Error", Messagebox.OK, Messagebox.ERROR);
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
		//SECTION WORK1 FocusedEvent
		try {
			Sheet sheet = event.getSheet();
			lastRow = event.getRow();
			lastCol = event.getColumn();
			
			Cell cell = Utils.getCell(sheet, lastRow, lastCol);

			//reading the text from cell to formula bar
			CellReference cr = new CellReference(lastRow, lastCol, false, false);
			String posStr = cr.formatAsString(); 
			lbpos.setLabel(posStr);
			focusPosition.setRawValue(posStr);
			focusPosition.setSelectedItem(lbpos);
			String editText = Utils.getEditText(cell);
			formulaEditbox.setValue(cell == null ? "" : (editText == null ? ""	: editText));
			
			if(cell != null){
				CellStyle cs = cell.getCellStyle();
				if (cs != null) {
				/*	int type=cell.getFormat().getFormatType();
					String typeName=null;
					if(type==0)
						typeName="general";
					if(type==1)
						typeName="text";
					if(type==2)
						typeName="number";
					if(type==3)
						typeName="datetime";
					if(type==16)
						typeName="control";
					Label lab = (Label) getFellow("formatTypeLabel");
					lab.setValue(" "+typeName);
					lab = (Label) getFellow("formatCodeLabel");
					if(cell.getFormat()!=null)
						lab.setValue(" "+cell.getFormat().getFormatCodes());
					else
						lab.setValue(" null");
					lab = (Label) getFellow("editTextLabel");
					lab.setValue(" "+cell.getEditText());
					lab = (Label) getFellow("resultLabel");
					if(cell.getResult()!=null)
						lab.setValue(" "+cell.getResult().toString());
					else
						lab.setValue(" null");
					lab = (Label) getFellow("valueLabel");
					if(cell.getValue()!=null)
						lab.setValue(" "+cell.getValue().toString());
					else
						lab.setValue(" null");
					lab = (Label) getFellow("textLabel");
					lab.setValue(" "+(String)cell.getText());
					
					Format format = cell.getFormat();
					if(format!=null){
						if(format.getBorderTop()!=null){
							lab = (Label) getFellow("top1Label");
							lab.setValue(format.getBorderTop().getBorderLineStyle().toString());
							lab = (Label) getFellow("top2Label");
							lab.setValue(format.getBorderTop().getBorderColor().toString());
						}
						if(format.getBorderRight()!=null){
							lab = (Label) getFellow("right1Label");
							lab.setValue(format.getBorderRight().getBorderLineStyle().toString());
							lab = (Label) getFellow("right2Label");
							lab.setValue(format.getBorderRight().getBorderColor().toString());
						}
						if(format.getBorderBottom()!=null){
							lab = (Label) getFellow("bottom1Label");
							lab.setValue(format.getBorderBottom().getBorderLineStyle().toString());
							lab = (Label) getFellow("bottom2Label");
							lab.setValue(format.getBorderBottom().getBorderColor().toString());
						}
						if(format.getBorderLeft()!=null){
							lab = (Label) getFellow("left1Label");
							lab.setValue(format.getBorderLeft().getBorderLineStyle().toString());
							lab = (Label) getFellow("left2Label");
							lab.setValue(format.getBorderLeft().getBorderColor().toString());
						}
						
						
						
					}*/
					//p("\nFormat Type:"+typeName);
					//p("EditText:"+cell.getEditText());
					//p("Result:"+cell.getResult());
					//p("Value:"+cell.getValue());
					//p("Text:"+cell.getText());
				}
			}
				
			//set to default setting
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
				
			//read format from cell and assign it to toolbar
			if(cell!=null){
				CellStyle cs = cell.getCellStyle();
				if(cs != null){
					int fontidx  = cs.getFontIndex();
					Font font = book.getFontAt((short)fontidx);
					
					fontFamilyCombobox.setText(font.getFontName());
					fontSizeCombobox.setText(Integer.toString(font.getFontHeightInPoints()));

					//font bold & italic
					isBold = font.getBoldweight() == Font.BOLDWEIGHT_BOLD;
					isItalic = font.getItalic();
					if(isBold)
						boldBtn.setClass("toolIcon clicked");
					
					if(isItalic)
						italicBtn.setClass("toolIcon clicked");

					//font underline 
					if(font.getUnderline()!= Font.U_NONE){
						isUnderline=true;
						underlineBtn.setClass("toolIcon clicked");
					}else{
						isUnderline=false;
					}

					//font stikethrough
					if(font.getStrikeout()){
						isStrikethrough=true;
						strikethroughBtn.setClass("toolIcon clicked");
					}else{
						isStrikethrough=false;
						
					}

					//align
					int align = cs.getAlignment();
					switch(align) {
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

					//font color
					String color = BookHelper.getFontColor(book, font);
					if (color != null && !color.equals(BookHelper.AUTO_COLOR)) {
						fontColorBtn.setColor(color);
					}
					
					//bg color					
					final int fillColorIdx = cs.getFillForegroundColor();
					String fcolor = BookHelper.indexToRGB(book, fillColorIdx);
					if (fcolor != null && !fcolor.equals(BookHelper.AUTO_COLOR)) {
						backgroundColorBtn.setColor(fcolor);
						backgroundColorMenu.setContent("#color=" + fcolor);
					}
					
					
					//merge cell
					//cannot find a way to know the cell is merged or not
					/*Toolbarbutton mergeCellBtn = (Toolbarbutton)getFellow("mergeCellBtn");
					isMergeCell=format.
					if(isWrapText){
						wrapTextBtn.setClass("toolIcon clicked");
					}else{
						wrapTextBtn.setClass("toolIcon");
					}*/
					
					//wrap text
					isWrapText=cs.getWrapText();
					if(isWrapText){
						wrapTextBtn.setClass("toolIcon clicked");
					}

					//FontStyle preFontStyle=format.getFontStyle();
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
		} 
	}
	

	//this is auto pushCellState in spreadsheet itself
	void onStopEditingEvent(StopEditingEvent evt){
		try{			
			//the formula bar input
			formulaEditbox.setValue((String)evt.getEditingValue());

			//to notify all widgets there is a cell changed
			for(int i=0; i<chartKey; i++){
				//p("chartKey="+chartKey+" chartWin"+i);
				try{
					Window win = (Window) mainWin.getFellow("chartWin" + i);
					if (win != null) {
						Chart myChart = (Chart) win.getFellow("myChart");
						CellEvent event = 
							new StopEditingEvent(org.zkoss.zss.ui.event.Events.ON_STOP_EDITING, myChart, evt.getSheet(), evt.getRow(), evt.getColumn(), (String) evt.getData());
						org.zkoss.zk.ui.event.Events.postEvent(event);
					}
				}catch(Exception e){
					//the chart may be deleted
					//e.printStackTrace();
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} 
	}

	void onEditboxEditingEvent(EditboxEditingEvent evt){
		try{			
			formulaEditbox.setValue((String)evt.getEditingValue());
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	void doCellChange(String value) {
		if (lastRow == -1) {
			return;
		}
		Sheet sheet = (Sheet) book.getSheetAt(0);
		//Cell cell = sheet.getCell(lastRow, lastCol);
		Utils.setCellValue(sheet, lastRow, lastCol, value);
		//if (cell == null) {
			//sheet.setCellValue(lastRow, lastCol, "");
			//cell = (Cell) sheet.getCell(lastRow, lastCol);
		//}
		//cell.setEditText(value);
	}

	public void SelectionEvent(CellSelectionEvent event) {
		//gonna try to print the current cell is merged or not
		int left = spreadsheet.getSelection().getLeft();
		int right = spreadsheet.getSelection().getRight();
		int top = spreadsheet.getSelection().getTop();
		int bottom = spreadsheet.getSelection().getBottom();

		isMergeCell=false;
		Sheet sheet=spreadsheet.getSelectedSheet();
		MergeMatrixHelper mmhelper = spreadsheet.getMergeMatrixHelper(sheet);
		for (final Iterator iter = mmhelper.getRanges().iterator(); iter.hasNext();) {
			MergedRect block = (MergedRect) iter.next();
			int bt = block.getTop();
			int bl = block.getLeft();
			int bb = block.getBottom();
			int br = block.getRight();
			if(left <= bl && top <= bt && right >= br && bottom >= bb){
				isMergeCell=true;
				break;
			}
		}
		
		Window win= (Window)mainWin.getFellow("fastIconContextmenu");
		//win.doPopup();
		Toolbarbutton mergeCellBtn2 = (Toolbarbutton)win.getFellow("_mergeCellBtn");
		if(isMergeCell){
			mergeCellBtn.setClass("toolIcon clicked");
			mergeCellBtn2.setClass("toolIcon clicked");
		}else{
			mergeCellBtn.setClass("toolIcon");
			mergeCellBtn2.setClass("toolIcon");
		}
	}
	
	//SECTION CtrlKeys
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
		char c=(char)event.getKeyCode();
		//p(""+event.getKeyCode()+c+" "+event.isCtrlKey());
		try {
			if(46==event.getKeyCode()){//delete
				if(event.isCtrlKey())
					onClearStyle((ForwardEvent)null);
				else
					onClearContent((ForwardEvent)null);
				return;
			}

			if (false == event.isCtrlKey()) {
				p("no ctrl");
				return;
			}
			
			p("Ctrl+"+c);
			
			switch (c) {
			case 'X':
				onEditCut();
				break;
			case 'C':
				onEditCopy();
				break;
			case 'V':
				onEditPaste();
				break;
			case 'D':
				onClearContent((ForwardEvent)null);
				break;
			case 'E'://for testing 
				spreadsheet.setCellFocus(new Position(2,2));
				spreadsheet.setSelection(new Rect(2,2,4,4));
				break;
			case 'S':
				fileh.dispatcher("save");
				break;
//TODO undo				
/*			case 'Z':
				spreadsheet.undo();
				break;
*/				
//TODO redo
/*			case 'Y':	
				spreadsheet.redo();
				break;
*/			case 'O':	
				fileh.selectDispatcher("open");
				break;
			case 'F':	

				break;
			case 'B':	
				onFontBoldClick();
				break;
			case 'I':	
				onFontItalicClick();
				break;
			case 'U':	
				onFontUnderlineClick();
				break;
			default:
				p("ctrl ...?");
			return;
			}


		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	//SECTION FILE MENU
	public void onFileSelect(ForwardEvent event) {
		fileh.selectDispatcher((String) event.getData());
	}
	public void onCellMenu(ForwardEvent event) {
		cellmh.dispatcher((String) event.getData());
	}
	public void onFileOK(ForwardEvent event) {
		fileh.dispatcher((String) event.getData());
	}
	public void onFileNew(ForwardEvent event){
		openNewFileInFS();
	}
	public void onFileSaveClose(ForwardEvent event){
		fileh.selectDispatcher("save");
		openNewFileInFS();
	}

	public void onOpenFile(ForwardEvent event){
		Listbox flo_files = (Listbox) Path.getComponent("//p1/mainWin/fileOpenWin/flo_files");
		String filename=null;
		if(flo_files.getSelectedItem()!=null){
			filename = flo_files.getSelectedItem().getLabel();
			p(filename);
			openFileInFS(filename);
		}
		
		Window fileOpenWin = (Window) Path.getComponent("//p1/mainWin/fileOpenWin");
		fileOpenWin.setVisible(false);
		//spreadsheet.conCurrentEditing(this);
		spreadsheet.invalidate();
	}
	
	public void onDeleteFile(ForwardEvent event){
		Listbox fld_files = (Listbox) Path.getComponent("//p1/mainWin/fileDeleteWin/fld_files");
		
		String filename = fld_files.getSelectedItem().getLabel();
		fileh.deleteFile(filename);

		Window fileDeleteWin = (Window) Path.getComponent("//p1/mainWin/fileDeleteWin");
		fileDeleteWin.setVisible(false);	
		
		if(filename.equals(spreadsheet.getSrc()))
			onFileNew(null);
	}

	public void onExportFile(ForwardEvent event){
		Listbox fle_files = (Listbox) Path.getComponent("//p1/mainWin/fileExportWin/fle_files");
		
		String filename = fle_files.getSelectedItem().getLabel();
		exportFile(filename);

		Window fileExportWin = (Window) Path.getComponent("//p1/mainWin/fileExportWin");
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
	public void onImportAndOpenFile(ForwardEvent event){
		openFileInFS(onImportFile(null));
	}
	public String onImportFile(ForwardEvent event) {
		String metaStr=null;
		String filename=null;
		try {
			Object ss = Fileupload.get();
			if(ss!=null){
				InputStream iStream = ((Media) ss).getStreamData();
				filename=((Media) ss).getName();
				long millis=System.currentTimeMillis();
				String hashFilename=millis+".xls";

				Files.copy(new File(fileh.xlsDir+hashFilename), iStream);

				FileOutputStream metaStream = new FileOutputStream(fileh.xlsDir+"metaFile",true);

				metaStr= new String(millis+"\n"+filename+"\n"+hashFilename+"\n");
				//System.out.println(fileh.xlsDir);
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
		
		File file=new File(fileh.xlsDir+"Untitled");
		
		try {
			openSpreadsheetFromStream(new FileInputStream(file), "Untitled");
			Label label = (Label) Path.getComponent("//p1/mainWin/openingFilename");
			label.setValue("Untitled");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
	
	public void exportFile(String filename) {// current or other files(stack_level=0)
//TODO
throw new UiException("export file not implmented yet");
/*
		// current editing file
		System.out.println("exportFile filename: " + filename + ", spreadsheet source: " + spreadsheet.getSrc());
		if (filename.equals(spreadsheet.getSrc())) {
			System.out.println("exportFile if: ");
			ByteArrayOutputStream baoStream = new ByteArrayOutputStream();
			ExcelExporter exporter = new ExcelExporter();
			exporter.exports(spreadsheet.getBook(), baoStream);
			byte[] bin = baoStream.toByteArray();
			Filedownload.save(new ByteArrayInputStream(bin), "application/vnd.ms-excel", filename);
		} else {// other files
			System.out.println("exportFile else");
			HashMap hm = fileh.readMetafile();
			Object objs[] = (Object[]) hm.get(filename);
			if (objs != null) {
				String hashFilename = fileh.xlsDir + ((String) objs[1]);
				System.out.println("hashFilename: " + hashFilename);
				FileInputStream iStream = null;
				try {
					iStream = new FileInputStream(hashFilename);
				} catch (FileNotFoundException e) {
					e.printStackTrace();
				}
				Filedownload.save(iStream, "application/vnd.ms-excel", filename);
			}
		}
*/	}

	public void onPrint(ForwardEvent event){

		String printKey=""+System.currentTimeMillis();
		Session session	=	Executions.getCurrent().getDesktop().getSession();
		session.setAttribute("zssFromHi"+printKey,spreadsheet);

		Window win = (Window) mainWin.getFellow("menuPrintWin");
		Button printBtn=(Button) win.getFellow("printBtn");
		printBtn.setHref("print.zul?printKey="+printKey);
		printBtn.setTooltip("print.zul?printKey="+printKey);

		win.setPosition("parent");
		win.setTop("100px");
		win.setLeft("100px");
		win.doPopup();
	}
	
	public void onRevision(ForwardEvent event){
		reloadRevisionMenu();
		Window win = (Window) mainWin.getFellow("revisionWin");
		win.setPosition("parent");
		win.setLeft("250px");
		win.setTop("250px");
		win.doPopup();
	}
	//SECTION EDIT MENU
	public void onEditUndo(ForwardEvent event){
//TODO undo		
		throw new UiException("Undo not implmented yet");
//		spreadsheet.undo();
	}
	
	public void onEditRedo(ForwardEvent event){
//TODO redo		
		throw new UiException("Redo not implemented yet");
//		spreadsheet.redo();
	}
	
	public void onEditCut(ForwardEvent event){
		onEditCut();
	}
	
	public void onEditCopy(ForwardEvent event){
		onEditCopy();
	}
	
	public void onEditPaste(ForwardEvent event){
		onEditPaste();
	}
	
	public void onEditCut(){
		try {
			isCut=true;
			remSheet = spreadsheet.getSelectedSheet();
			spreadsheet.setHighlight(spreadsheet.getSelection());
		} catch(Exception e) {

		}
	}
	
	public void onEditCopy(){
		try {
			isCut=false;
			remSheet = spreadsheet.getSelectedSheet();
			spreadsheet.setHighlight(spreadsheet.getSelection());	
		} catch(Exception e) {

		}
	}
	
	private boolean isOverlapMergedCell(Sheet sheet, int top, int left, int bottom, int right) {
		MergeMatrixHelper mmhelper = spreadsheet.getMergeMatrixHelper(sheet);
		for(final Iterator iter = mmhelper.getRanges().iterator(); iter.hasNext();) {
			final MergedRect block = (MergedRect) iter.next();
			int bl = block.getLeft();
			int br = block.getRight();
			int bt = block.getTop();
			int bb = block.getBottom();
			if (bt <= bottom && bl <= right 
				&& br >= left && bb >= top)
				return true;
		}
		return false;
	}
	
	public void onEditPaste(){
		if(spreadsheet.getHighlight()==null)
			return;
		
		int srcLeft = spreadsheet.getHighlight().getLeft();
		int srcRight = spreadsheet.getHighlight().getRight();
		int srcTop = spreadsheet.getHighlight().getTop();
		int srcBottom = spreadsheet.getHighlight().getBottom();

		int dstLeft = spreadsheet.getSelection().getLeft();
		int dstTop = spreadsheet.getSelection().getTop();
		int dstRight = dstLeft + (srcRight - srcLeft);
		int dstBottom = dstTop + (srcBottom - srcTop);
		
		
		// using RangeSimple
		Sheet sheet = spreadsheet.getSelectedSheet();
		Range srcRect = Utils.getRange(remSheet, srcTop, srcLeft, srcBottom, srcRight);
		Range dstRect = Utils.getRange(sheet, dstTop, dstLeft, dstBottom, dstRight);
		
		if(isOverlapMergedCell(sheet, dstTop, dstLeft, dstBottom, dstRight)) { 
			try {
				Messagebox.show("cannot change part of merged cell in destination region");
				return;
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
			
		int maxRows = spreadsheet.getMaxrows();
		int maxCols = spreadsheet.getMaxcolumns();
		if (dstRight + 1 > maxCols) {
			spreadsheet.setMaxcolumns(dstRight + 1);
		}
		if (dstBottom + 1 > maxRows) {
			spreadsheet.setMaxrows(dstBottom + 1);
		}
		
		//save the state
		Sheet[] sheets = new Sheet[2];
		sheets[0]=remSheet;
		sheets[1]=spreadsheet.getSelectedSheet();
		
		Rect rectArray[] = new Rect[2];
		rectArray[0] = new Rect(srcLeft, srcTop, srcRight, srcBottom);
		rectArray[1] = new Rect(dstLeft, dstTop, dstRight, dstBottom);
		
		int srcSheetIndex=spreadsheet.indexOfSheet(sheets[0]);
		int dstSheetIndex=spreadsheet.indexOfSheet(sheets[1]);
		
//TODO undo/redo		
//		spreadsheet.pushCellState(sheets, rectArray);
		
		srcRect.copy(dstRect);
//TODO undo/redo
//		((CellState)spreadsheet.peekUndoState()).pasteCellState(sheets[1], new Position(dstTop,dstLeft));
		boolean isSameSheet=(srcSheetIndex==dstSheetIndex);
//TODO cut
/*
		if (isCut==true) {//clean the src
			Cell cell=null;
			
			if(isSameSheet){
				for(int row=srcTop; row<=srcBottom;row++)
					for(int col=srcLeft; col<=srcRight;col++){
						if(row>=dstTop && row<=dstBottom && col>=dstLeft && col<=dstRight)
							continue;
						cell = Utils.getCell(remSheet, row, col);
						if(cell!=null){
							cell.setEditText(null);
							cell.setFormat(new FormatImpl());
						}
					}
				remSheet.unmergeCells(srcLeft, srcTop, srcRight, srcBottom);
				spreadsheet.setHighlight(null);
			}
		}
*/		
		if(isSameSheet && !(dstLeft>srcRight || dstRight<srcLeft || dstTop>srcBottom || dstBottom<srcTop)){
			spreadsheet.setHighlight(null);
		}
	}
	
	public void onClearContent(ForwardEvent event){
//TODO undo/redo		
//		spreadsheet.pushCellState();
		onClearContent();
	}
	
	public void onClearContent(){
		try {
			int left = spreadsheet.getSelection().getLeft();
			int right = spreadsheet.getSelection().getRight();
			int top = spreadsheet.getSelection().getTop();
			int bottom = spreadsheet.getSelection().getBottom();
			Sheet sheet = spreadsheet.getSelectedSheet();
			Cell cell=null;
			for(int row=top; row<=bottom;row++)
				for(int col=left; col<=right;col++){
					cell = Utils.getCell(sheet,row,col);
					if(cell!=null){
						Utils.setEditText(cell, null);
					}
				}
		} catch (Exception e) {
			e.printStackTrace();
		}	
	}
	
	public void onClearStyle(ForwardEvent event){
//TODO undo/redo		
//		spreadsheet.pushCellState();
		onClearStyle();
	}
	
	public void onClearStyle(){
		try {
			int left = spreadsheet.getSelection().getLeft();
			int right = spreadsheet.getSelection().getRight();
			int top = spreadsheet.getSelection().getTop();
			int bottom = spreadsheet.getSelection().getBottom();
			Sheet sheet = spreadsheet.getSelectedSheet();
			Cell cell=null;
			for(int row=top; row<=bottom;row++) {
				for(int col=left; col<=right;col++){
					cell = Utils.getCell(sheet,row,col);
					if(cell!=null){
						cell.setCellStyle(null);
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}	
	}

	public void onClearBoth(ForwardEvent event){
//TODO undo/redo		
//		spreadsheet.pushCellState();
		onClearStyle();
		onClearContent();
	}
	
	public void onDeleteRows(ForwardEvent event){
		try {
			Sheet sheet=spreadsheet.getSelectedSheet();
			Rect rect = spreadsheet.getSelection();
			int top=rect.getTop();
			int left=rect.getLeft();
			int bottom = rect.getBottom();
			if(top==0 && bottom==spreadsheet.getMaxrows()-1){
				try {
					Messagebox.show("cannot delete all Rows");
					return;
				} catch (InterruptedException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
//TODO undo/redo			
//			spreadsheet.pushDeleteRowColState(-1, top, -1, bottom);
			Utils.deleteRows(sheet, top, bottom);
			spreadsheet.setSelection(rect);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void onDeleteColumns(ForwardEvent event){
		try {
			Sheet sheet=spreadsheet.getSelectedSheet();
			Rect rect = spreadsheet.getSelection();
			int left = rect.getLeft();
			int right = rect.getRight();
			if(left==0 && right==spreadsheet.getMaxcolumns()-1){
				try {
					Messagebox.show("cannot delete all Columns");
					return;
				} catch (InterruptedException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
//TODO undo/redo			
//			spreadsheet.pushDeleteRowColState(left, -1, right, -1);
			Utils.deleteColumns(sheet, left, right);
			spreadsheet.setSelection(rect);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	//SECTION VIEW MENU
	public void onViewFormulaBar(ForwardEvent event) {
		if (formulaBar.getHeight() != "0px") {
			formulaBar.setSize("0px");
			topToolbars.setHeight("53px");
		} else {
			topToolbars.setHeight("76px");
			formulaBar.setSize("23px");
		}
	}

	public void onViewFreezeRows(ForwardEvent event) {
		spreadsheet.setRowfreeze(Integer.parseInt((String)event.getData())-1);
	}
	
	public void onViewFreezeCols(ForwardEvent event) {
		spreadsheet.setColumnfreeze(Integer.parseInt((String)event.getData())-1);
	}
	
	public void onViewUnfreezeRows(ForwardEvent event){
		spreadsheet.setRowfreeze(-1);
	}
	
	public void onViewUnfreezeCols(ForwardEvent event){
		spreadsheet.setColumnfreeze(-1);
	}

	//SECTION FORMAT MENU
	public void onFormatNumber(ForwardEvent event){
		Window win = (Window) mainWin.getFellow("formatNumberWin");
		win.setPosition("parent");
		win.setLeft("170px");
		win.setTop("24px");	
		win.doPopup();//Modal();
	}
	
	/**
	 * seems no need
	 * @param event
	 */
	/*
	public void onFormatBackgroundClick(ForwardEvent event){
		colorSelectorTarget="background";
		onColorPicker(event);
	}
	*/
	
	public void onTestMenuAlignment() {
		try {
			Messagebox.show("called FormatAlignment onOK");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	//SECTION INSERT MENU
	public void onInsertFormula(ForwardEvent event){
		Window win = (Window) mainWin.getFellow("formulaCategory");
		win.setPosition("parent");
		win.setLeft("220px");
		win.setTop("24px");	
		win.doPopup();
	}

	public void onInsertRows(ForwardEvent event){
		try {
			final Rect rect = spreadsheet.getSelection();
			final int top=rect.getTop();
			final int bottom = rect.getBottom();
//TODO undo/redo			
//			spreadsheet.pushInsertRowColState(-1, top, -1, bottom);
			Utils.insertRows(spreadsheet.getSelectedSheet(), top, bottom);
			rect.setTop(bottom+1);
			rect.setBottom(bottom+bottom-top+1);
			spreadsheet.setCellFocus(new Position(rect.getTop(),rect.getLeft()));
			spreadsheet.setSelection(rect);
			
		} catch (Exception e) {
			e.printStackTrace();
		} 
	}

	public void onInsertColumns(ForwardEvent event){
		try {
			Rect rect = spreadsheet.getSelection();
			int left = rect.getLeft();
			int right = rect.getRight();

//TODO undo/redo
//			spreadsheet.pushInsertRowColState(left, -1, right, -1);
			Utils.insertColumns(spreadsheet.getSelectedSheet(), left, right);
			rect.setLeft(right+1);
			rect.setRight(right+right-left+1);
			spreadsheet.setCellFocus(new Position(rect.getTop(),rect.getLeft()));
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
	

	//SECTION HELP MENU
	public void onHelpCheatsheet(ForwardEvent event){
		Window win = (Window) mainWin.getFellow("cheatsheet");
		win.setPosition("parent");
		win.setLeft("327px");
		win.setTop("124px");	
		win.doPopup();
	}
	
	//SECTION Formula bar
	public void onFormulaBarBtn(ForwardEvent event){
		event_x = 207;
		event_y = 101;
		onFormulaPopup();
	}

	public void onFormulaListOK() {
		Combobox formulaList = (Combobox) Path.getComponent("//p1/mainWin/formulaList");

		int left = spreadsheet.getSelection().getLeft();
		int top = spreadsheet.getSelection().getTop();
		p("wh left= " + String.valueOf(left));
		p("wh top= " + String.valueOf(top));
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
	
	public void onEditingFormulaInput(Event event){
		if (currentEditcell == null)
			return;
		if (currentEditcell.getCellType() == Cell.CELL_TYPE_FORMULA && currentEditcell.getCellFormula() != null)
			return;
		
		Utils.setEditText(currentEditcell, ((InputEvent) event).getValue());
	}

	private void changeFont(Sheet sheet, int top, int left, int bottom, int right) {
		//font
		String name = (String) this.fontFamilyCombobox.getSelectedItem().getLabel();
		String fontSizeStr = this.fontSizeCombobox.getSelectedItem().getLabel();
		short fontHeight = Short.parseShort(fontSizeStr);
		short boldWeight = isBold ? Font.BOLDWEIGHT_BOLD : Font.BOLDWEIGHT_NORMAL;
		boolean italic = isItalic;
		byte underline = isUnderline ? Font.U_SINGLE : Font.U_NONE;
		boolean strikeout = isStrikethrough;
		String colorrgb = fontColorBtn.getColor();
		short typeOffset = Font.SS_NONE;
		short color = BookHelper.rgbToIndex(book, colorrgb);
		
		//alignment
		
		boolean everchanged = false;
		Font font = book.findFont(boldWeight, color, fontHeight, name, italic, strikeout, typeOffset, underline);
		if (font != null) {
			for (int row = top; row <= bottom; row++)
				for (int col = left; col <= right; col++) {
					Cell cell = Utils.getOrCreateCell(sheet, row, col);
					CellStyle cs = cell.getCellStyle();
					final Font srcFont = book.getFontAt(cs.getFontIndex());
					if (srcFont.equals(font)) //same font, no need to change it
						continue;
					//create a new CellStyle and use the new font
					CellStyle ncs = book.createCellStyle();
					ncs.cloneStyleFrom(cs);
					ncs.setFont(font);
					cell.setCellStyle(ncs);
					everchanged = true;
				}
		}
		//TODO notify cell change
		//if (everchanged) notify
	}
	
	//SECTION FastIcon Toolbar
	public void onFontFamilySelect(ForwardEvent event) {
//TODO undo/redo		
//		spreadsheet.pushCellState();
		try {
			Combobox fontFamily = null;
			String font;
			Window win = (Window) mainWin.getFellow("fastIconContextmenu");
			if (win.isVisible()) {
				win.setVisible(false);
				fontFamily = (Combobox) win.getFellow("_fontFamilyCombobox");
			} else {
				fontFamily = this.fontFamilyCombobox;
			}
			if (event.getData() == null)
				font = fontFamily.getSelectedItem().getLabel();
			else
				font = (String) event.getData();

			this.fontFamilyCombobox.setSelectedIndex((fontFamily
					.getSelectedIndex()));

			int left = spreadsheet.getSelection().getLeft();
			int right = spreadsheet.getSelection().getRight();
			int top = spreadsheet.getSelection().getTop();
			int bottom = spreadsheet.getSelection().getBottom();
			Sheet sheet = spreadsheet.getSelectedSheet();
			changeFont(sheet, top, left, bottom, right);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void onFontSizeSelect(ForwardEvent event) {
//TODO undo/redo		
//		spreadsheet.pushCellState();
		try {
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

			int left = spreadsheet.getSelection().getLeft();
			int right = spreadsheet.getSelection().getRight();
			int top = spreadsheet.getSelection().getTop();
			int bottom = spreadsheet.getSelection().getBottom();
			Sheet sheet = spreadsheet.getSelectedSheet();
			changeFont(sheet, top, left, bottom, right);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void onFontBoldClick(ForwardEvent event){
		onFontBoldClick();
	}
	public void onFontBoldClick(){
//TODO undo/redo	
//		spreadsheet.pushCellState();
		try{
			mainWin.getFellow("fastIconContextmenu").setVisible(false);	
			isBold = !isBold;
			if (isBold) {
				boldBtn.setClass("toolIcon clicked");
			} else {
				boldBtn.setClass("toolIcon");
			}

			int left = spreadsheet.getSelection().getLeft();
			int right = spreadsheet.getSelection().getRight();
			int top = spreadsheet.getSelection().getTop();
			int bottom = spreadsheet.getSelection().getBottom();
			Sheet sheet = spreadsheet.getSelectedSheet();
			changeFont(sheet, top, left, bottom, right);
		}catch(Exception e){
			e.printStackTrace();
		} 
	}

	public void onFontItalicClick(ForwardEvent event){
		onFontItalicClick();
	}
	public void onFontItalicClick(){
//TODO undo/redo		
//		spreadsheet.pushCellState();
		try{
			mainWin.getFellow("fastIconContextmenu").setVisible(false);
			isItalic = !isItalic;
			if (isItalic) {
				italicBtn.setClass("toolIcon clicked");
			} else {
				italicBtn.setClass("toolIcon");
			}

			int left = spreadsheet.getSelection().getLeft();
			int right = spreadsheet.getSelection().getRight();
			int top = spreadsheet.getSelection().getTop();
			int bottom = spreadsheet.getSelection().getBottom();
			Sheet sheet = spreadsheet.getSelectedSheet();
			changeFont(sheet, top, left, bottom, right);
		}catch(Exception e){
			e.printStackTrace();
			p("onFontFamilySelect");
		} 
	}

	public void onFontUnderlineClick(ForwardEvent event){
		onFontUnderlineClick();
	}
	public void onFontUnderlineClick(){
//TODO undo/redo		
//		spreadsheet.pushCellState();
		try{
			isUnderline=!isUnderline;
			if(isUnderline){
				underlineBtn.setClass("toolIcon clicked");
			}else{
				underlineBtn.setClass("toolIcon");
			}

			int left = spreadsheet.getSelection().getLeft();
			int right = spreadsheet.getSelection().getRight();
			int top = spreadsheet.getSelection().getTop();
			int bottom = spreadsheet.getSelection().getBottom();
			Sheet sheet = spreadsheet.getSelectedSheet();
			changeFont(sheet, top, left, bottom, right);
		}catch(Exception e){
			e.printStackTrace();
		} 
	}

	public void onFontStrikethroughClick(ForwardEvent event){
//TODO undo/redo		
//		spreadsheet.pushCellState();
		try{
			isStrikethrough=!isStrikethrough;
			if(isStrikethrough){
				strikethroughBtn.setClass("toolIcon clicked");
			}else{
				strikethroughBtn.setClass("toolIcon");
			}

			int left = spreadsheet.getSelection().getLeft();
			int right = spreadsheet.getSelection().getRight();
			int top = spreadsheet.getSelection().getTop();
			int bottom = spreadsheet.getSelection().getBottom();
			Sheet sheet = spreadsheet.getSelectedSheet();
			changeFont(sheet, top, left, bottom, right);
		}catch(Exception e){
			e.printStackTrace();
		} 
	}

	public void onAlignHorizontalClick(ForwardEvent event){
//TODO undo/redo
//		spreadsheet.pushCellState();
		
		mainWin.getFellow("fastIconContextmenu").setVisible(false);
		String alignStr=(String)event.getData();

		int left = spreadsheet.getSelection().getLeft();
		int right = spreadsheet.getSelection().getRight();
		int top = spreadsheet.getSelection().getTop();
		int bottom = spreadsheet.getSelection().getBottom();
		Sheet sheet = spreadsheet.getSelectedSheet();

		short align = CellStyle.ALIGN_GENERAL;

		if(alignStr.equals("left")){
			alignLeftBtn.setClass("toolIcon clicked");
			align=CellStyle.ALIGN_LEFT;
		}else
			alignLeftBtn.setClass("toolIcon");

		if(alignStr.equals("center")){
			alignCenterBtn.setClass("toolIcon clicked");
			align=CellStyle.ALIGN_CENTER;
		}else
			alignCenterBtn.setClass("toolIcon");

		if(alignStr.equals("right")){
			alignRightBtn.setClass("toolIcon clicked");
			align=CellStyle.ALIGN_RIGHT;
		}else
			alignRightBtn.setClass("toolIcon");

		boolean everchanged = false;
		for (int row = top; row <= bottom; row++) {
			for (int col = left; col <= right; col++) {
				Cell cell = Utils.getOrCreateCell(sheet, row, col);
				CellStyle cs = cell.getCellStyle();
				final short srcAlign = cs.getAlignment();
				if (srcAlign == align) //same color, no need to change it
					continue;
				//create a new CellStyle and use the new color
				CellStyle ncs = book.createCellStyle();
				ncs.cloneStyleFrom(cs);
				ncs.setAlignment(align);
				cell.setCellStyle(ncs);
				everchanged = true;
			}
		}
		
		//TODO everchanged, notify change
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

		int left = spreadsheet.getSelection().getLeft();
		int right = spreadsheet.getSelection().getRight();
		int top = spreadsheet.getSelection().getTop();
		int bottom = spreadsheet.getSelection().getBottom();
		Sheet sheet = spreadsheet.getSelectedSheet();
		changeFont(sheet, top, left, bottom, right);
	}
	
	public void setBackgroundColor(String color) {
		if (color == null)
			return;
		
		int left = spreadsheet.getSelection().getLeft();
		int right = spreadsheet.getSelection().getRight();
		int top = spreadsheet.getSelection().getTop();
		int bottom = spreadsheet.getSelection().getBottom();
		Sheet sheet = spreadsheet.getSelectedSheet();
		short colorIndex = BookHelper.rgbToIndex(book, color);
		backgroundColorBtn.setColor(color);
		
		boolean everchanged = false;
		for (int row = top; row <= bottom; row++)
			for (int col = left; col <= right; col++) {
				Cell cell = Utils.getOrCreateCell(sheet, row, col);
				CellStyle cs = cell.getCellStyle();
				final short srcColor = cs.getFillForegroundColor();
				if (srcColor == colorIndex) //same color, no need to change it
					continue;
				//create a new CellStyle and use the new color
				CellStyle ncs = book.createCellStyle();
				ncs.cloneStyleFrom(cs);
				ncs.setFillForegroundColor(colorIndex);
				cell.setCellStyle(ncs);
				everchanged = true;
			}
	}

	/**
	 *
	 * @param event
	 */
	public void onWrapTextClick(ForwardEvent event){
//TODO remove me, wrap text
throw new UiException("wrap text is implmented yet");
/*
		spreadsheet.pushCellState();
		try{
			isWrapText=!isWrapText;
			if(isWrapText){
				wrapTextBtn.setClass("toolIcon clicked");
			}else{
				wrapTextBtn.setClass("toolIcon");
			}

			int left = spreadsheet.getSelection().getLeft();
			int right = spreadsheet.getSelection().getRight();
			int top = spreadsheet.getSelection().getTop();
			int bottom = spreadsheet.getSelection().getBottom();
			Sheet sheet = spreadsheet.getSelectedSheet();
			for (int row = top; row <= bottom; row++) 
				for (int col = left; col <= right; col++){
					Styles.setTextWrap(sheet, row, col, isWrapText);		
				}
		}catch(Exception e){
			e.printStackTrace();
		}
*/
	}

	public void onMergeCellClick(ForwardEvent event){
		try{
			mainWin.getFellow("fastIconContextmenu").setVisible(false);

			//isMergeCell should read from the cell
			//and search all over the cell if any one is merged then unmerge all
			isMergeCell=!isMergeCell;
			if(isMergeCell){
				mergeCellBtn.setClass("toolIcon clicked");
			}else{
				mergeCellBtn.setClass("toolIcon");
			}

			if(isMergeCell){
//TODO: undo/redo				
//				spreadsheet.pushCellState();
				Rect sel = spreadsheet.getSelection();
				if(sel.getLeft()-sel.getRight()==0){
					mergeCellBtn.setClass("toolIcon");
				}else{
					//TODO: true mean merge horizontal only.(UI cannot handle merge vertically yet)
					Utils.mergeCells(spreadsheet.getSelectedSheet(), sel.getTop(), sel.getLeft(), sel.getBottom(), sel.getRight(), true);
				}
				spreadsheet.focus();
			}else{
//TODO: undo/redo				
//				spreadsheet.pushCellState();
				Rect sel = spreadsheet.getSelection();
				Utils.unmergeCells(spreadsheet.getSelectedSheet(), sel.getTop(), sel.getLeft(), sel.getBottom(), sel.getRight());
				spreadsheet.focus();
			}
		}catch(Exception e){
			e.printStackTrace();
		}
	}

	private void sortAscending() {
		Utils.sort(spreadsheet.getSelectedSheet(), spreadsheet.getSelection(), null, null, null, false, false, false);
	}

	private void sortDescending() {
		Utils.sort(spreadsheet.getSelectedSheet(), spreadsheet.getSelection(), null, new boolean[]{true}, null, false, false, false);
	}

	public void onClick$sortAscending() {
		sortAscending();
	}

	public void onClick$sortDescending() {
		sortDescending();
	}

	public void onClick$sortAscendingMenu() {
		sortAscending();
	}

	public void onClick$sortDescendingMenu() {
		sortDescending();
	}

	public void onClick$customSort() {
		HashMap arg = new HashMap();
		arg.put("spreadsheet", spreadsheet);
		Executions.createComponents("/menus/sort/customSort.zul", mainWin, arg);
	}
	
	public void onClick$copySelection() {
		Rect sel = spreadsheet.getSelection();
		if (sel.getBottom() >= spreadsheet.getMaxrows())
			sel.setBottom(spreadsheet.getMaxrows() - 1);
		if (sel.getRight() >= spreadsheet.getMaxcolumns())
			sel.setRight(spreadsheet.getMaxcolumns() - 1);
		spreadsheet.setHighlight(sel);
	}
	
	public void onClick$pasteSelection() {
		if (spreadsheet.getHighlight() == null)
			return;

		Utils.pasteSpecial(spreadsheet.getSelectedSheet(), 
				spreadsheet.getHighlight(), 
				spreadsheet.getSelection().getTop(), spreadsheet.getSelection().getLeft(), 
				Range.PASTE_ALL, Range.PASTEOP_NONE, 
				false, false);
		spreadsheet.setHighlight(null);
	}

	public void onClick$pasteSpecial() {
		HashMap arg = new HashMap();
		arg.put("spreadsheet", spreadsheet);
		if (spreadsheet.getHighlight() != null)
			Executions.createComponents("/menus/paste/pasteSpecial.zul", mainWin, arg);
		else {
			//To do: paste from Clipboard Dialog (pasteSpecialFromClipboard.zul)
			try {
				Messagebox.show("Please select copy area");
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
		}
	}

	public void onBorderClick(ForwardEvent event){
		try {	
			Window win = (Window) mainWin.getFellow("borderSelector");
			win.setPosition("parent");
			if(event.getData()!=null && ((String)event.getData()).equals("toolbar")){
				win.setLeft("668px");
				win.setTop("52px");	
			}else{
				mainWin.getFellow("fastIconContextmenu").setVisible(false);
				((Window)mainWin.getFellow("fastIconContextmenu")).doPopup();
				win.setLeft(80 + event_x + "px");
				win.setTop(-108 + event_y + "px");
			}
			win.doPopup();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void onBorderSelector(ForwardEvent event){
//TODO: undo/redo
//spreadsheet.pushCellState();
		String presetStyle=(String)event.getData();
		//p(presetStyle);
		//not permit to changing color and style
		final BorderStyle lineStyle = BorderStyle.MEDIUM;
		final String color = "#000000";

		Window win = (Window) mainWin.getFellow("borderSelector");
		win.setVisible(false);
		int lCol = spreadsheet.getSelection().getLeft();
		int rCol = spreadsheet.getSelection().getRight();
		int tRow = spreadsheet.getSelection().getTop();
		int bRow = spreadsheet.getSelection().getBottom();
		Sheet sheet = spreadsheet.getSelectedSheet();
		Range rng = Utils.getRange(sheet, tRow, lCol, bRow, rCol);
		
		//check the is merge cell this time for right border
		if (presetStyle.equals("None")) {
			rng.setBorders(BookHelper.BORDER_FULL, BorderStyle.NONE, color);
		} else if (presetStyle.equals("Outline")) {
			rng.setBorders(BookHelper.BORDER_OUTLINE, lineStyle, color);
		} else if (presetStyle.equals("Inside")) {
			rng.setBorders(BookHelper.BORDER_INSIDE, lineStyle, color);
		} else if (presetStyle.equals("Full")) {
			rng.setBorders(BookHelper.BORDER_FULL, lineStyle, color);
		} else if (presetStyle.equals("Top")) {
			final Range tRowRng = Utils.getRange(sheet, tRow, lCol, tRow, rCol);
			tRowRng.setBorders(BookHelper.BORDER_EDGE_TOP, lineStyle, color);
		} else if (presetStyle.equals("Bottom")) {
			final Range bRowRng = Utils.getRange(sheet, bRow, lCol, bRow, rCol);
			bRowRng.setBorders(BookHelper.BORDER_EDGE_BOTTOM, lineStyle, color);
		} else if (presetStyle.equals("Left")) {
			final Range lColRng = Utils.getRange(sheet, tRow, lCol, bRow, lCol);
			lColRng.setBorders(BookHelper.BORDER_EDGE_LEFT, lineStyle, color);
		} else if (presetStyle.equals("Right")) {
			final Range rColRng = Utils.getRange(sheet, tRow, rCol, bRow, rCol);
			rColRng.setBorders(BookHelper.BORDER_EDGE_RIGHT, lineStyle, color);
		}
	}

	public void onFormatOK(ForwardEvent event) {
		//p("formatOk"+(String)event.getData());
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
	
	
	//SECTION Utility Functions
	public void p(String s){
		//System.out.println(spreadsheet.getEditorName()+": "+s);
	}

	public void onFormulaPopup() {
		try {
			Window win = (Window) mainWin.getFellow("formulaCategory");
			win.setPosition("parent");
			win.setLeft(event_x + "px");
			win.setTop(event_y + "px");
			win.doPopup();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void onFormatPopup(){
		Window win = (Window) mainWin.getFellow("formatNumberWin");
		try {
			//the set position only work at the second time?
			win.setPosition("parent");
			win.setLeft(event_x + "px");
			win.setTop(event_y + "px");	
			win.doPopup();//Modal();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Write a chart widget ??
	 * @param evt
	 */
	public void onClick$insertPieChart(Event evt) {
//TODO remove me, insert PieChart
throw new UiException("insert pie chart is not implmented yet");
/*
		Component cmp = execution.createComponents("piechart.zul", mainWin, null);
		if (cmp != null) {
			System.out.println("create cmp");
		}
		Chart chart = (Chart)cmp.getFellow("mychart");
		ChartWidget wgt = new ChartWidget();
		wgt.setChart(chart);

		int col = spreadsheet.getSelection().getLeft();
		int row = spreadsheet.getSelection().getTop();
		wgt.setRow(row);
		wgt.setColumn(col);
		SpreadsheetCtrl ctrl = (SpreadsheetCtrl) spreadsheet.getExtraCtrl();
		ctrl.addWidget(wgt);
*/	}
	

	public void onInsertChart(ForwardEvent event){
//TODO remove me, insert chart
throw new UiException("insert chart is not implmented yet");		
/*		final int left = spreadsheet.getSelection().getLeft();
		final int right = spreadsheet.getSelection().getRight();
		final int top = spreadsheet.getSelection().getTop();
		final int bottom = spreadsheet.getSelection().getBottom();
		Sheet sheet = spreadsheet.getSelectedSheet();

		final String modelType=(String)event.getData();
		
		//BarModel model2=new SimpleBarModel();
		StringBuffer strBuf = new StringBuffer();
		strBuf.append("<window title=\"test\" width=\"420px\" id=\"chartWin"+(chartKey++)+"\" border=\"normal\">");
		strBuf.append("<chart id=\"myChart\" width=\"400px\" height=\"200px\" type=\""+modelType+"\" threeD=\"true\" fgAlpha=\"128\">");
		strBuf.append("<zscript>");
		p(modelType);
		if(modelType.equals("pie")){
			strBuf.append("myChart.setModel(new SimplePieModel());");
		}
		if(modelType.equals("bar")){
			strBuf.append("myChart.setModel(new SimpleCategoryModel());");
		}
		strBuf.append("</zscript>");
		strBuf.append("</chart>");
		strBuf.append("</window>");
		Window win= (Window) Executions.createComponentsDirectly(strBuf.toString(), null, mainWin, null);
		Chart myChart = (Chart) win.getFellow("myChart");
		ChartModel model=myChart.getModel();
		for(int row=top; row<=bottom; row++){
			Cell cellName=sheet.getCell(row, left);
			String name;
			if(cellName==null || cellName.getText()==null)
				name="";
			else
				name=cellName.getText();
			for(int col=left+1; col<=right; col++){
				Cell cellValue=sheet.getCell(row, col);
				double value;
				if(cellValue==null || cellValue.getResult()==null)
					value=0;
				else{
					try{
						value=((Double)cellValue.getResult()).doubleValue();
					}catch(Exception e){
						value=0;
					}

					//p(name);
					//p(""+value);
					if(modelType.equals("pie")){
						((PieModel)model).setValue((Comparable)name,new Double(value));
					}
					if(modelType.equals("bar")){
						((CategoryModel)model).setValue((Comparable)name, (Comparable)new Long(col-left), new Double(value));
					}
				}
			}
		}
		if(modelType.equals("pie"))
			myChart.setModel((PieModel)model);
		if(modelType.equals("bar"))
			myChart.setModel((CategoryModel)model);

		win.doOverlapped();
		//win.setSizable(true);
		win.setClosable(true);
		//win.setAttribute(SpreadsheetCtrl.CHILD_PASSING_KEY,"");
		//spreadsheet.appendChild(win);
		mainWin.appendChild(win);

		win.setLeft("100px");
		win.setTop("200px");
		win.doOverlapped();
		//spreadsheet.smartUpdate("appendWidget",win.getUuid());




		//Add EventListener for updating chart
		myChart.addEventListener(Events.ON_STOP_EDITING,
				new EventListener() {
			public void onEvent(Event event) throws Exception {
				int modRow=((CellEvent)event).getRow();
				int modCol=((CellEvent)event).getColumn();
				if(left<=modCol && modCol<=right && top<=modRow && modRow<=bottom){
					//p("ChartValue Editing:"+(String)event.getData());
					//p("row: "+top+" "+bottom);
					//p("col: "+left+" "+right);
					//p("event:"+modelType);
					Sheet sheet=((CellEvent)event).getSheet();
					Chart myChart=(Chart)event.getTarget();
					ChartModel model = myChart.getModel();
					if(modelType.equals("pie"))
						((PieModel)model).clear();
					if(modelType.equals("bar"))
						((CategoryModel)model).clear();
					
					for(int row=top; row<=bottom; row++){
						Cell cellName=sheet.getCell(row, left);
						String name;
						if(cellName==null || cellName.getText()==null)
							name="";
						else
							name=cellName.getText();

						for(int col=left+1; col<=right; col++){
							Cell cellValue=sheet.getCell(row, col);
							double value;
							if(cellValue==null || cellValue.getResult()==null)
								value=0;
							else{
								try{
									value=((Double)cellValue.getResult()).doubleValue();
								}catch(Exception e){
									value=0;
								}

								//p(name);
								//p(""+value);
								if(modelType.equals("pie")){
									((PieModel)model).setValue((Comparable)name,new Double(value));
								}
								if(modelType.equals("bar")){
									((CategoryModel)model).setValue((Comparable)name, (Comparable)new Long(col-left), new Double(value));
								}
							}
						}
					}
					if(modelType.equals("pie"))
						myChart.setModel((PieModel)model);
					if(modelType.equals("bar"))
						myChart.setModel((CategoryModel)model);

				}

			}
		});

		win.addEventListener(org.zkoss.zk.ui.event.Events.ON_CLICK, 
				new EventListener(){
			
			public void onEvent(Event event) throws Exception {
				Rect rect=new Rect();
				rect.setBottom(bottom);
				rect.setTop(top);
				rect.setRight(right);
				rect.setLeft(left);
				spreadsheet.setSelection(rect);
			}
		});
*/
	}

	
	public void reloadRevisionMenu(){

		try
        {	
			//read the metafile
			BufferedReader bReader;
			bReader = new BufferedReader(new FileReader(fileh.xlsDir+"metaFile"));
			String line=null;
			Stack stack = new Stack();
			
			String timeStr,filename,hashFilename;
			while(true){
				timeStr=bReader.readLine();
				if(timeStr==null){
					break;
				}
				filename=bReader.readLine();
				if(filename==null){
					System.out.println("Warning: filename cannot read from metaFile");
					break;
				}
				hashFilename=bReader.readLine();
				if(hashFilename==null){
					System.out.println("Warning: hashFilename cannot read from metaFile");
					break;
				}
				if(spreadsheet.getSrc().equals(filename)){
					stack.add(timeStr);
					stack.add(filename);
					stack.add(hashFilename);
				}
			}
			bReader.close();

			//put read data to Menu
            org.zkoss.zul.Row newRow;
			Rows revisionRows = (Rows)mainWin.getFellow("revisionWin").getFellow("revisionRows");
			List rowsChildList = revisionRows.getChildren();
			while(!rowsChildList.isEmpty())
				revisionRows.removeChild((Component)rowsChildList.get(0));
			while(!stack.isEmpty()){
				
				hashFilename=(String) stack.pop();
				filename=(String) stack.pop();
				timeStr=(String) stack.pop();
				String dateStr = new Date(Long.parseLong(timeStr)).toString();
            	
            	newRow = new org.zkoss.zul.Row(); 
            	Radio radio = new Radio();
            	radio.setAttribute("value", hashFilename);
            	
            	
            	if(hashFilename.equals(this.hashFilename)){
            		newRow.appendChild(new Label("current"));
            		newRow.setStyle("background:rgb(250,230,180) none");
            	}else{
            		newRow.appendChild(radio);
            		//p("set hashFilename: "+hashFilename);
            	}
                newRow.appendChild(new Label(dateStr));
                newRow.appendChild(new Label("test name"));
                newRow.appendChild(new Label("no comment"));
            	
                revisionRows.appendChild(newRow);
                //p(""+radio.getAttribute("value"));
            }
            //close the jdbc connection
            
            
        }
        catch (Exception ex)
        {
            p("exception: "+ex.getMessage());
            
        }
	}
	

	public void onRevisionOK(ForwardEvent event){
		String filename=null;
		
		Rows revisionRows = (Rows)mainWin.getFellow("revisionWin").getFellow("revisionRows");
		List rowList = revisionRows.getChildren();
		for(int i = 0; i<rowList.size(); i++){
			Component tmpComponent=((Component)rowList.get(i)).getFirstChild();
			if(tmpComponent instanceof Radio && ((Radio)tmpComponent).isChecked()){
				filename=(String)((Radio)tmpComponent).getAttribute("value");
				//p("onRevision: "+filename);
			}
		}
		
		openFileInFS(filename);
		
		Window revisionWin = (Window) Path.getComponent("//p1/mainWin/revisionWin");
		revisionWin.setVisible(false);
		spreadsheet.invalidate();
		
		/**
		 * 
		 */
        //spreadsheet.notifyRevision();
           
	}

	
	public void openSpreadsheetFromStream(InputStream iStream, String src){
		
		spreadsheet.setBookFromStream(iStream, src);
    	
		book=spreadsheet.getBook();
		sheetTBClean();
		sheetTBInit();
		sheetTB.addEventListener(org.zkoss.zk.ui.event.Events.ON_SELECT,
				new EventListener() {
			public void onEvent(Event event) throws Exception {
				onTabboxSelectEvent((SelectEvent) event);
			}
		});
		
		spreadsheet.setSelectedSheet(((Tab)sheetTB.getTabs().getFirstChild()).getLabel());
		spreadsheet.invalidate();
	}
	
	private Connection getDBConnection(){
		try {
			return DriverManager.getConnection("jdbc:mysql://localhost:3306/zss","root","rootzk");
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	public void onNotImplement(ForwardEvent event){
		try {
			Messagebox.show("Not implement yet");
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
	}
	
	public void onDeleteSheet(){
		//TODO Check if there is the last sheet
		if(spreadsheet.getBook().getNumberOfSheets()==1){
			try {
				Messagebox.show("cannot remove last sheet, but you could insert a new sheet, then remove current sheet  ");
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			return;
		}
			
		//TODO show the hint "Are you sure?"
		String name = spreadsheet.getSelectedSheet().getSheetName();
		int buttons = Messagebox.YES+Messagebox.NO;
		int result = -1;
		try {
			result=Messagebox.show("Do you really want to delete selected sheet \""+name+"\", those data will be deleted permanently", "",buttons, "");
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
		//clean the related state in the undo/redo stack
//TODO undo/redo		
//		spreadsheet.cleanRelatedState(spreadsheet.getSheet(index));
		
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
}