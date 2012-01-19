/* Spreadsheet.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Thu May 17 14:41:33     2007, Created by tomyeh
		Dec 19 10:10:10 2007, modify by Dennis Chen.
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
 */
package org.zkoss.zss.ui;

import static org.zkoss.zss.ui.fn.UtilFns.getCellFormatText;
import static org.zkoss.zss.ui.fn.UtilFns.getCelltext;
import static org.zkoss.zss.ui.fn.UtilFns.getEdittext;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.Serializable;
import java.net.URL;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.zkoss.json.JSONArray;
import org.zkoss.json.JSONObject;
import org.zkoss.lang.Classes;
import org.zkoss.lang.Library;
import org.zkoss.lang.Objects;
import org.zkoss.lang.Strings;
import org.zkoss.poi.ss.SpreadsheetVersion;
import org.zkoss.poi.ss.formula.FormulaParseException;
import org.zkoss.poi.ss.usermodel.AutoFilter;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Chart;
import org.zkoss.poi.ss.usermodel.DataValidation;
import org.zkoss.poi.ss.usermodel.DataValidation.ErrorStyle;
import org.zkoss.poi.ss.usermodel.FilterColumn;
import org.zkoss.poi.ss.usermodel.Picture;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.usermodel.ZssChartX;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.poi.ss.util.CellRangeAddressList;
import org.zkoss.poi.ss.util.CellReference;
import org.zkoss.util.logging.Log;
import org.zkoss.util.media.AMedia;
import org.zkoss.util.media.Media;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.xel.Function;
import org.zkoss.xel.FunctionMapper;
import org.zkoss.xel.VariableResolver;
import org.zkoss.xel.XelException;
import org.zkoss.xml.HTMLs;
import org.zkoss.zk.au.AuRequest;
import org.zkoss.zk.au.AuResponse;
import org.zkoss.zk.au.out.AuInvoke;
import org.zkoss.zk.ui.AbstractComponent;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Desktop;
import org.zkoss.zk.ui.Execution;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Page;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.WebApp;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.EventQueues;
import org.zkoss.zk.ui.ext.render.DynamicMedia;
import org.zkoss.zk.ui.sys.ContentRenderer;
import org.zkoss.zk.ui.sys.JavaScriptValue;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.engine.event.EventDispatchListener;
import org.zkoss.zss.engine.event.SSDataEvent;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Importer;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.model.impl.BookCtrl;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.model.impl.ExcelImporter;
import org.zkoss.zss.ui.au.in.BlockSyncCommand;
import org.zkoss.zss.ui.au.in.CellFetchCommand;
import org.zkoss.zss.ui.au.in.CellFocusedCommand;
import org.zkoss.zss.ui.au.in.CellMouseCommand;
import org.zkoss.zss.ui.au.in.CellSelectionCommand;
import org.zkoss.zss.ui.au.in.Command;
import org.zkoss.zss.ui.au.in.EditboxEditingCommand;
import org.zkoss.zss.ui.au.in.FilterCommand;
import org.zkoss.zss.ui.au.in.HeaderCommand;
import org.zkoss.zss.ui.au.in.HeaderMouseCommand;
import org.zkoss.zss.ui.au.in.MoveWidgetCommand;
import org.zkoss.zss.ui.au.in.SelectionChangeCommand;
import org.zkoss.zss.ui.au.in.StartEditingCommand;
import org.zkoss.zss.ui.au.in.StopEditingCommand;
import org.zkoss.zss.ui.au.in.WidgetCtrlKeyCommand;
import org.zkoss.zss.ui.au.out.AuCellFocus;
import org.zkoss.zss.ui.au.out.AuCellFocusTo;
import org.zkoss.zss.ui.au.out.AuDataUpdate;
import org.zkoss.zss.ui.au.out.AuHighlight;
import org.zkoss.zss.ui.au.out.AuInsertRowColumn;
import org.zkoss.zss.ui.au.out.AuMergeCell;
import org.zkoss.zss.ui.au.out.AuRemoveRowColumn;
import org.zkoss.zss.ui.au.out.AuRetrieveFocus;
import org.zkoss.zss.ui.au.out.AuSelection;
import org.zkoss.zss.ui.au.out.AuUpdateData;
import org.zkoss.zss.ui.event.CellEvent;
import org.zkoss.zss.ui.event.CellSelectionEvent;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.event.HyperlinkEvent;
import org.zkoss.zss.ui.event.SheetDeleteEvent;
import org.zkoss.zss.ui.event.StartEditingEvent;
import org.zkoss.zss.ui.event.StopEditingEvent;
import org.zkoss.zss.ui.impl.CellFormatHelper;
import org.zkoss.zss.ui.impl.HeaderPositionHelper;
import org.zkoss.zss.ui.impl.HeaderPositionHelper.HeaderPositionInfo;
import org.zkoss.zss.ui.impl.JSONObj;
import org.zkoss.zss.ui.impl.MergeAggregation;
import org.zkoss.zss.ui.impl.MergeAggregation.MergeIndex;
import org.zkoss.zss.ui.impl.MergeMatrixHelper;
import org.zkoss.zss.ui.impl.MergedRect;
import org.zkoss.zss.ui.impl.SequenceId;
import org.zkoss.zss.ui.impl.StringAggregation;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zss.ui.sys.SpreadsheetCtrl;
import org.zkoss.zss.ui.sys.SpreadsheetCtrl.CellAttribute;
import org.zkoss.zss.ui.sys.SpreadsheetInCtrl;
import org.zkoss.zss.ui.sys.SpreadsheetOutCtrl;
import org.zkoss.zss.ui.sys.WidgetHandler;
import org.zkoss.zss.ui.sys.WidgetLoader;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.impl.XulElement;


/**
 * Spreadsheet is a rich ZK Component to handle EXCEL like behavior, it reads
 * the data from a data model({@link Book}) and interact with this model by
 * event.<br/>
 * You can assign a Book by {@link #setBook(Book)} or just set the .xls file
 * location by {@link #setSrc(String)}. You also need to set two attributes to
 * restrict max rows and columns to show on client side by
 * {@link #setMaxrows(int)} and {@link #setMaxcolumns(int)}. <br/>
 * To use Spreadsheet in .zul file, just use <code>&lt;spreadsheet/&gt;</code>
 * tag like any other ZK Components.<br/>
 * An simplest example : <br/>
 * <span
 * style="font-family: courier new; font-size: 10pt;">&nbsp;&nbsp;&nbsp;&nbsp
 * ;<span style="color: rgb(0,128,128);">&lt;</span><span
 * style="color: rgb(63,127,127);">spreadsheet&nbsp;</span><span
 * style="color: rgb(127,0,127);">src</span>=<span
 * style="color: rgb(42,0,255);">
 * &quot;/WEB-INF/xls/my.xls&quot;&nbsp;</span><span
 * style="color: rgb(127,0,127);">maxrows</span>=<span
 * style="color: rgb(42,0,255);">&quot;300&quot;&nbsp;</span><span
 * style="color: rgb(127,0,127);">maxcolumns</span>=<span
 * style="color: rgb(42,0,255);">&quot;80&quot;&nbsp;<br />
 * &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><span
 * style="color: rgb(127,0,127);">height</span>=<span
 * style="color: rgb(42,0,255);">&quot;400px&quot;&nbsp;</span><span
 * style="color: rgb(127,0,127);">width</span>=<span
 * style="color: rgb(42,0,255);">&quot;90%&quot;&nbsp;</span><span
 * style="color: rgb(0,128,128);">/&gt;</span></span>
 * 
 * 
 * @author dennischen
 */

public class Spreadsheet extends XulElement implements Serializable {
	private static final Log log = Log.lookup(Spreadsheet.class);

	private static final long serialVersionUID = 1L;
	private static final String ROW_SIZE_HELPER_KEY = "_rowCellSize";
	private static final String COLUMN_SIZE_HELPER_KEY = "_colCellSize";
	private static final String MERGE_MATRIX_KEY = "_mergeRange";
	private static final String WIDGET_HANDLER = "org.zkoss.zss.ui.sys.WidgetHandler.class";
	private static final String WIDGET_LOADERS = "org.zkoss.zss.ui.sys.WidgetLoader.class";

	transient private Book _book; // the spreadsheet book

	private String _src; // the src to create an internal book
	transient private Importer _importer; // the spreadsheet importer
	private int _maxRows = 20; // how many row of this spreadsheet
	private int _maxColumns = 10; // how many column of this spreadsheet
	
	private int _preloadRowSize = -1; //the number of row to load when receiving the rendering request
	private int _preloadColumnSize = -1; //the number of column to load when receiving the rendering request
	
	private String _selectedSheetId;
	transient private Worksheet _selectedSheet;
	transient private String _selectedSheetName;

	private int _rowFreeze = -1; // how many fixed rows
	private boolean _rowFreezeset = false;
	private int _colFreeze = -1; // how many fixed columns
	private boolean _colFreezeset = false;
	private boolean _hideRowhead; // hide row head
	private boolean _hideColhead; // hide column head*/
	
	private boolean _hideGridlines; //hide gridlines
	private boolean _protectSheet;
	private boolean _showFormulabar;
	
	//TODO undo/redo
	//StateManager stateManager = new StateManager(this);
	private Map<String, Focus> _focuses = new HashMap<String, Focus>(20); //id -> Focus

	private Rect _focusRect = new Rect(0, 0, 0, 0);
	private Rect _selectionRect = new Rect(0, 0, 0, 0);
	private Rect _visibleRect = new Rect();
	private Rect _loadedRect = new Rect();
	private Rect _highlightRect = null;

	private WidgetHandler _widgetHandler;

	private List<WidgetLoader> _widgetLoaders;


	/**
	 * default row height when a sheet is empty
	 */
	private int _defaultRowHeight = 20;

	/**
	 * default column width when a sheet is empty
	 */
	private int _defaultColumnWidth = 64;

	/**
	 * dynamic css version
	 */
	private int _cssVersion = 0;

	/**
	 * width of left header
	 */
	private int _leftheadWidth = 36;

	/**
	 * height of top header
	 */
	private int _topheadHeight = 20;

	/**
	 * cell padding of each cell and header, both on left and right side.
	 */
	private int _cellpadding = 2;

	/**
	 * customized row and column names.
	 */
	private Map _columnTitles;
	private Map _rowTitles;

	private InnerDataListener _dataListener = new InnerDataListener();
	private InnerVariableResolver _variableResolver = new InnerVariableResolver();
	private InnerFunctionMapper _functionMapper = new InnerFunctionMapper();

	/**
	 * Server side customized column/row id, start with 1,3,5,7. If a
	 * column,which set from client side, the id will be 2,4,6 , check ss.js for
	 * detail
	 */
	private SequenceId _custColId = new SequenceId(-1, 2);
	private SequenceId _custRowId = new SequenceId(-1, 2);
	private SequenceId _updateCellId = new SequenceId(0, 1);// to handle batch
	private SequenceId _updateRangeId = new SequenceId(0, 1);
	private SequenceId _focusId = new SequenceId(0, 1);

	private String _userName;
	
	public Spreadsheet() {
		this.addEventListener("onStartEditingImpl", new EventListener() {
			public void onEvent(Event event) throws Exception {
				Object[] data = (Object[]) event.getData();
				processStartEditing((String) data[0],
						(StartEditingEvent) data[1], (String) data[2]);
			}
		});
		this.addEventListener("onStopEditingImpl", new EventListener() {
			public void onEvent(Event event) throws Exception {

				Object[] data = (Object[]) event.getData();
				processStopEditing((String) data[0], (StopEditingEvent) data[1], (String) data[2]);
			}
		});
	}

	/**
	 * Don't call this, the spreadsheet is not draggable.
	 * 
	 * @exception UnsupportedOperationException if this method is called.
	 */
	public void setDraggable(String draggable) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Returns the book model of this Spreadsheet. If you call this method at
	 * first time and the book has not assigned by {@link #setBook(Book)}, this
	 * will create a new model depends on src;
	 * 
	 * @return the book model of this spread sheet.
	 */
	public Book getBook() {
		if (_book == null) {
			if (_src == null) {
				return null;
			}
			try {
				Importer importer = _importer;
				if (importer == null) {
					importer = new ExcelImporter();
				}

				Book book = null;
				if (importer instanceof ExcelImporter) {
					URL url = null;
					
					if (_src.startsWith("/")) {// try to load by application
						// context.
						WebApp wapp = Executions.getCurrent().getDesktop().getWebApp();
						String path = wapp.getRealPath(_src);
						if (path != null) {
							File file = new File(path);
//							if (file.isDirectory())
//								throw new IllegalArgumentException("Your input source is a directory, not a vaild file");
							if (file.exists())
								url = file.toURI().toURL();
						} else
							url = wapp.getResource(_src); 
					}
					if (url == null) {// try to load from class loader
						url = new ClassLocator().getResource(_src);
					}
					if (url == null) {// try to load from file
						File f = new File(_src);
						if (f.exists()) {
							url = f.toURI().toURL();
						}
					}

					if (url == null) {
						throw new UiException("resource for " + _src + " not found.");
					}

					book = ((ExcelImporter) importer).importsFromURL(url);
				} else {
					book = importer.imports(_src);
				}

				initBook(book); //will set _book inside this method
			} catch (Exception ex) {
				throw UiException.Aide.wrap(ex);
			}
		}
		return _book;
	}

	/**
	 * Sets the book data model of this spread sheet.
	 * 
	 * @param book the book data model.
	 */
	public void setBook(Book book) {
		if (!Objects.equals(book, _book)) {
			initBook0(book);
			invalidate();
		}
	}
	
	private void initBook(Book book) {
		if (!Objects.equals(book, _book)) {
			initBook0(book);
		}
	}
	private Focus _focus;
	private void deleteFocus() {
		if (_selectedSheet != null && _focus != null) {
			final Range rng = Ranges.range(_selectedSheet);
			rng.notifyDeleteFriendFocus(_focus);
			((BookCtrl)_book).removeFocus(_focus);
			_focus = null;
		}
	}
	private void moveFocus() {
		if (_selectedSheet != null) {
			if (_focus == null) {
				_focus = newFocus();
				((BookCtrl)_book).addFocus(_focus);
			}
			final Range rng = Ranges.range(_selectedSheet);
			rng.notifyMoveFriendFocus(_focus);
		}
	}
	private EventListener _focusListener = null;
	private void doMoveSelfFocus(CellEvent event){
		syncEditorFocus();
		int row=event.getRow();
		int col=event.getColumn();
		_focus.row = row;
		_focus.col = col;
		moveFocus();
	}
	private void initBook0(Book book) {
		if (_book != null) {
			if (_focusListener != null)
				removeEventListener(Events.ON_CELL_FOUCSED, _focusListener);
			deleteFocus();
			_book.unsubscribe(_dataListener);
			_book.removeVariableResolver(_variableResolver);
			_book.removeFunctionMapper(_functionMapper);
			if (_focusListener != null)
				removeEventListener(Events.ON_CELL_FOUCSED, _focusListener);
		}
		
		 //Shall clean selected sheet before set new book (ZSS-75: set book null, cause NPE)
		if (_selectedSheet != null) {
			doSheetClean(_selectedSheet);
		}
		_selectedSheet = null;
		_selectedSheetId = null;
		_selectedSheetName = null;
		
		_book = book;
		if (_book != null) {
			_book.subscribe(_dataListener);
			_book.addVariableResolver(_variableResolver);
			_book.addFunctionMapper(_functionMapper);
			if (_focus == null) { //focus name default to Spreadsheet id
				_focus = newFocus();
				((BookCtrl)_book).addFocus(_focus);
			}
			if (EventQueues.APPLICATION.equals(_book.getShareScope()) || EventQueues.SESSION.equals(_book.getShareScope()) ) { //have to sync focus
				this.addEventListener(Events.ON_CELL_FOUCSED, _focusListener = new EventListener() {
					@Override
					public void onEvent(Event event) throws Exception {
						doMoveSelfFocus((CellEvent) event);
					}
				});
			}
		}
	}
	
	private Focus newFocus() {
		final String focusName = _userName == null ? ""+ (getId() == null ? getUuid() : getId()) : _userName;
		return new Focus(((BookCtrl)_book).nextFocusId(), focusName, "#000", 0, 0, this);
	}
	
	/**
	 * 
	 * @param is
	 * @param src
	 */
	public void setBookFromStream(InputStream is, String src){
		if(_selectedSheet!=null){
			doSheetClean(_selectedSheet);
		}
		_selectedSheet=null;
		_selectedSheetId=null;
		_selectedSheetName = null;
		//bug#315: freezed pane rows/columns don't work when setting Spreadsheet from Composer.
		//setRowfreeze(-1);
		//setColumnfreeze(-1);
		//setBook(null);
		_importer = new ExcelImporter();
		_src=src;
		final Book book = ((ExcelImporter)_importer).imports(is, src);
		initBook(book);
		invalidate();
	}

	/**
	 * Gets the selected sheet, the default selected sheet is first sheet.
	 * @return #{@link Worksheet}
	 */
	public Worksheet getSelectedSheet() {
		final Book book = getBook();
		if (book == null) {
			return null;
		}
		if (_selectedSheet == null) {
			if (_selectedSheetId == null) {
				if (book.getNumberOfSheets() == 0)
					throw new UiException("sheet size of given book is zero");
				_selectedSheet = (Worksheet) book.getSheetAt(0);
				_selectedSheetId = Utils.getSheetUuid(_selectedSheet);
			} else {
				_selectedSheet = Utils.getSheetByUuid(_book, _selectedSheetId);
				if (_selectedSheet == null) {
					throw new UiException("can not find sheet by id : "	+ _selectedSheetId);
				}
			}
			doSheetSelected(_selectedSheet);
		}
		return _selectedSheet;
	}

	public Worksheet getSheet(int index){
		final Book book = getBook();
		return book != null ? (Worksheet)book.getSheetAt(index) : null;
	}
	
	public int indexOfSheet(Worksheet sheet){
		final Book book = getBook();
		return book != null ? book.getSheetIndex(sheet) : -1;
	}
	
	private String getSelectedSheetId() {
		if (_selectedSheetId == null) {
			getSelectedSheet();
		}
		return _selectedSheetId;
	}

	/**
	 * Returns the src location of book model. This src is used by the specified
	 * importer to create the book data model of this spread sheet.
	 * 
	 * @return the src location
	 */
	public String getSrc() {
		return _src;
	}


	/**
	 * Sets the src location of the book data model to be imported into
	 * spreadsheet. A specified importer ({@link #getImporter}) will use this
	 * src to create the book data model.
	 * 
	 * @param src  the book src location
	 */
	public void setSrc(String src) {
		if (!Objects.equals(_src, src)) {
			_src = src;
			setBook(null);
			invalidate();
		}
	}

	/**
	 * it will change the spreadsheet src name as src
	 * and book name as src => will not reload the book
	 * @param src
	 */
	public void setSrcName(String src) {
		/**
		 * TODO model integration
		 */
		/*
		BookImpl book = (BookImpl) this.getBook();
		if (book != null)
			book.setName(src);
		_src = src;
		*/
		final Book book = (Book) this.getBook();
		if (book != null)
			_src = src;
	}
	
	/**
	 * Gets the importer that import the file in the specified src (
	 * {@link #getSrc}) to {@link Book} data model. The default importer is
	 * {@link ExcelImporter}.
	 * 
	 * @return the importer
	 */
	public Importer getImporter() {
		return _importer;
	}

	/**
	 * Sets the importer for import the book data model from a specified src.
	 * 
	 * @param importer the importer to import a spread sheet file from a document
	 * format (e.g. an Excel file) by the specified src (@link
	 * #setSrc(). The default importer is {@link ExcelImporter}.
	 */
	public void setImporter(Importer importer) {
		if (!Objects.equals(importer, _importer)) {
			_importer = importer;
			setBook(null);
		}
	}

	/**
	 * Sets the selected sheet by a name
	 * @param name	the name of spreadsheet to be selected.
	 */
	public void setSelectedSheet(String name) {
		final Book book = getBook();
		if (book == null) {
			return;
		}
		
		//Note. check whether if the sheet has remove or not
		if (_selectedSheet != null && book.getSheetIndex(_selectedSheet) == -1) {
			doSheetClean(_selectedSheet);
			_selectedSheet = null;
			_selectedSheetId = null;
			_selectedSheetName = null;
		}

		if (_selectedSheet == null || !_selectedSheet.getSheetName().equals(name)) {
			Worksheet sheet = book.getWorksheet(name);
			if (sheet == null) {
				throw new UiException("No such sheet : " + name);
			}

			if (_selectedSheet != null)
				doSheetClean(_selectedSheet);
			_selectedSheet = sheet;
			_selectedSheetId = Utils.getSheetUuid(_selectedSheet);
			invalidate();
			// the call of onSheetSelected must after invalidate,
			// because i must let invalidate clean lastcellblock first
			doSheetSelected(_selectedSheet);
		}
	}

	/**
	 * Returns the maximum visible number of rows of this spreadsheet. You can assign
	 * new number by calling {@link #setMaxrows(int)}.
	 * 
	 * @return the maximum visible number of rows.
	 */
	public int getMaxrows() {
		return _maxRows;
	}

	/**
	 * Sets the maximum visible number of rows of this spreadsheet. For example, if you set
	 * this parameter to 40, it will allow showing only row 0 to 39. The minimal value of max number of rows
	 * must large than 0; <br/>
	 * Default : 40.
	 * 
	 * @param maxrows  the maximum visible number of rows
	 */
	public void setMaxrows(int maxrows) {
		if (maxrows < 1) {
			throw new UiException("maxrow must be greater than 0: " + maxrows);
		}

		if (_maxRows != maxrows) {
			_maxRows = maxrows;
			if (_rowFreeze >= _maxRows) {
				_rowFreeze = _maxRows - 1;
			}
			
			smartUpdate("maxRow", getMaxrowInJSON());
		}
	}
	
	public void setPreloadRowSize(int size) {
		if (_preloadRowSize != size) {
			_preloadRowSize = size <= 0 ? 0 : size;
			smartUpdate("preloadRowSize", _preloadRowSize);
		}
	}
	
	public int getPreloadRowSize() {
		return _preloadRowSize;
	}

	private String getMaxrowInJSON() {
		JSONObj result = new JSONObj();
		result.setData("maxrow", _maxRows);
		//issue #98: Freeze area is not rendered if it is on the second sheet
		result.setData("rowfreeze", getRowfreeze());
		return result.toString();
	}
	
	/**
	 * Returns the maximum visible number of columns of this spreadsheet. You can assign
	 * new numbers by calling {@link #setMaxcolumns(int)}.
	 * 
	 * @return the maximum visible number of columns 
	 */
	public int getMaxcolumns() {
		return _maxColumns;
	}

	/**
	 * Sets the maximum visible number of columns of this spreadsheet. For example, if you
	 * set this parameter to 40, it will allow showing only column 0 to column 39. the minimal value of
	 * max number of columns must large than 0;
	 * 
	 * @param maxcols  the maximum visible number of columns
	 */
	public void setMaxcolumns(int maxcols) {
		if (maxcols < 1) {
			throw new UiException("maxcolumn must be greater than 0: " + maxcols);
		}

		if (_maxColumns != maxcols) {
			_maxColumns = maxcols;

			if (_colFreeze >= _maxColumns) {
				_colFreeze = _maxColumns - 1;
			}
			smartUpdate("maxColumn", getMaxcolumnsInJSON());
		}
	}
	
	public void setPreloadColumnSize(int size) {
		if (_preloadColumnSize != size) {
			_preloadColumnSize = size < 0 ? 0 : size;
			smartUpdate("preloadColSize", _preloadRowSize);
		}
	}
	
	public int getPreloadColumnSize() {
		return _preloadColumnSize;
	}

	private String getMaxcolumnsInJSON() {
		JSONObj result = new JSONObj();
		result.setData("maxcol", _maxColumns);
		//issue #98: Freeze area is not rendered if it is on the second sheet
		result.setData("colfreeze", getColumnfreeze());
		return result.toString();
	}

	/**
	 * Returns the row freeze index of this spreadsheet, zero base. A minus
	 * value means don't freeze row. If you setRowfreeze by
	 * {@link #setRowfreeze(int)}, then it always return the value that you set.
	 * Else if there is a selected sheet, it get row freeze from selected sheet
	 * and keep it(which means, it will always return this value if you does set
	 * it again), then return it. Default : -1
	 * 
	 * @return the row freeze of this spreadsheet or selected sheet.
	 */
	public int getRowfreeze() {
		if (_rowFreezeset)
			return _rowFreeze;
		final Worksheet sheet = getSelectedSheet();
		if (sheet != null) {
			 if (BookHelper.isFreezePane(sheet)) { //issue #103: Freeze row/column is not correctly interpreted
				 _rowFreeze = BookHelper.getRowFreeze(sheet) - 1;
			 }
			_rowFreezeset = true;
		}
		return _rowFreeze;
	}

	/**
	 * Sets the row freeze of this spreadsheet, sets row freeze on this
	 * component has higher priority then row freeze on sheet
	 * 
	 * @param rowfreeze row index
	 */
	public void setRowfreeze(int rowfreeze) {
		_rowFreezeset = true;
		if (rowfreeze < 0) {
			rowfreeze = -1;
		}
		if (_rowFreeze != rowfreeze) {
			_rowFreeze = rowfreeze;
			invalidate();
		}
	}

	/**
	 * Returns the column freeze index of this spreadsheet, zero base. A minus
	 * value means don't freeze column. If you set column freeze by
	 * {@link #setColumnfreeze(int)}, then it always return the value that you
	 * set. Else if there is a selected sheet, it get column freeze from
	 * selected sheet and keep it(which means, it will always return this value
	 * if you does set it again), then return it. Default : -1
	 * 
	 * @return the column freeze of this spreadsheet or selected sheet.
	 */
	public int getColumnfreeze() {
		if (_colFreezeset)
			return _colFreeze;
		Worksheet sheet = getSelectedSheet();
		if (sheet != null) {
			 if (BookHelper.isFreezePane(sheet)) {//issue #103: Freeze row/column is not correctly interpreted
				 _colFreeze = BookHelper.getColumnFreeze(sheet) - 1;
			 }
			_colFreezeset = true;
		}
		return _colFreeze;
	}

	/**
	 * Sets the column freeze of this spreadsheet, sets row freeze on this
	 * component has higher priority then row freeze on sheet
	 * 
	 * @param columnfreeze  column index
	 */
	public void setColumnfreeze(int columnfreeze) {
		_colFreezeset = true;
		if (columnfreeze < 0) {
			columnfreeze = -1;
		}
		if (_colFreeze != columnfreeze) {
			_colFreeze = columnfreeze;
			invalidate();
		}
	}

	/**
	 * Returns true if hide row head of this spread sheet; default to false.
	 * @return true if hide row head of this spread sheet; default to false.
	 */
	public boolean isHiderowhead() {
		return _hideRowhead;
	}

	/**
	 * Sets true to hide the row head of this spread sheet.
	 * @param hide true to hide the row head of this spread sheet.
	 */
	public void setHiderowhead(boolean hide) {
		// TODO
		if (_hideRowhead != hide) {
			_hideRowhead = hide;
			invalidate();
			/**
			 * rename. hiderowhead -> rowHeadHidden, not implement yet
			 */
			//smartUpdate("rowHeadHidden", _hideRowhead);
		}
	}

	/**
	 * Returns true if hide column head of this spread sheet; default to false.
	 * @return true if hide column head of this spread sheet; default to false.
	 */
	public boolean isHidecolumnhead() {
		return _hideColhead;
	}

	/**
	 * Sets true to hide the column head of this spread sheet.
	 * @param hide true to hide the row head of this spread sheet.
	 */
	public void setHidecolumnhead(boolean hide) {
		// TODO
		if (_hideColhead != hide) {
			_hideColhead = hide;
			invalidate();
			/**
			 * rename. hidecolhead ->columnHeadHidden, not implement yet
			 */
			// smartUpdate("z.hidecolhead", b);
			// smartUpdate("columnHeadHidden", _hideColhead);
		}
	}

	/**
	 * Sets the customized titles of column header.
	 * 
	 * @param titles a map for customized column titles, the key of column must be Integer object.
	 */
	public void setColumntitles(Map titles) {
		if (!Objects.equals(titles, _columnTitles)) {
			_columnTitles = titles;
			invalidate();
		}
	}

	/**
	 * Get the column titles Map object, modification of return object doesn't
	 * cause any update.
	 * 
	 * @return Map object of customized column names
	 */
	public Map getColumntitles() {
		return _columnTitles;
	}

	/**
	 * Sets the customized titles which split by ','. For example:
	 * "name,title,age" or "name,,age" , an empty string means use default
	 * column title. This method will split the input string to a Map with
	 * sequence index from 0, then call {@link #setColumntitles(Map)}<br/>
	 * 
	 * <p/>
	 * Note: this method will always invoke invalidate()
	 * 
	 * @param titles  the column titles
	 */
	public void setColumntitles(String titles) {
		String[] names = titles.split(",");
		Map map = new HashMap();
		for (int i = 0; i < names.length; i++) {
			if (names[i].length() > 0)
				map.put(Integer.valueOf(i), names[i]);
		}
		if (map.size() > 0)
			setColumntitles(map);

	}

	/**
	 * Sets the customized titles of row header.
	 * 
	 * @param titles map for customized column titles, the key of column must be Integer object.
	 */
	public void setRowtitles(Map titles) {
		if (!Objects.equals(titles, _rowTitles)) {
			_rowTitles = titles;
			invalidate();
		}
	}

	/**
	 * Get the row titles Map object, modification of the return object doesn't
	 * cause any update.
	 * 
	 * @return Map object of customized row names
	 */
	public Map getRowtitles() {
		return _rowTitles;
	}

	/**
	 * Sets the customized titles which split by ','. For example:
	 * "name,title,age" or "name,,age" , an empty string means use default row
	 * title. This method will split the input string to a Map with sequence
	 * index from 0, then call {@link #setRowtitles(Map)}<br/>
	 * 
	 * <p/>
	 * Note: this method will always invoke invalidate()
	 * 
	 * @param titles the row titles
	 */
	public void setRowtitles(String titles) {
		String[] names = titles.split(",");
		Map map = new HashMap();
		for (int i = 0; i < names.length; i++) {
			if (names[i].length() > 0)
				map.put(Integer.valueOf(i), names[i]);
		}
		if (map.size() > 0)
			setRowtitles(map);
	}

	/**
	 * Gets the default row height of the selected sheet
	 * @return default value depends on selected sheet
	 */
	public int getRowheight() {
		Worksheet sheet = getSelectedSheet();
		int rowHeight = sheet != null ? sheet.getDefaultRowHeight() : -1;

		return (rowHeight <= 0) ? _defaultRowHeight : Utils.twipToPx(rowHeight);
	}

	
	/**
	 * Sets the default row height of the selected sheet
	 * @param rowHeight the row height
	 */
	public void setRowheight(int rowHeight) {
		Worksheet sheet = getSelectedSheet();
		int rowHeightTwip = Utils.pxToTwip(rowHeight);
		int dh = sheet.getDefaultRowHeight();

		if (dh != rowHeightTwip) {
			sheet.setDefaultRowHeight((short)rowHeightTwip);
			invalidate();
			/**
			 * rename rowh -> rowHeight, not implement yet
			 */
			//smartUpdate("rowHeight", rowHeight);
		}
	}
	
	/**
	 * Gets the default column width of the selected sheet
	 * @return default value depends on selected sheet
	 */
	public int getColumnwidth() {
		final Worksheet sheet = getSelectedSheet();
		return Utils.getDefaultColumnWidthInPx(sheet);
	}

	/**
	 * Sets the default column width of the selected sheet
	 * @param columnWidth the default column width
	 */
	public void setColumnwidth(int columnWidth) {
		final Worksheet sheet = getSelectedSheet();
		int dw = sheet.getDefaultColumnWidth();
		if (dw != columnWidth) {
			sheet.setDefaultColumnWidth(columnWidth);
			
			invalidate();
			/**
			 * rename colw -> columnWidth , not implement yet
			 */
			//smartUpdate("columnWidth", columnWidth);
		}
	}
	
	private int getDefaultCharWidth() {
		final Worksheet sheet = getSelectedSheet();
		return Utils.getDefaultCharWidth(sheet);
	}

	/**
	 * Gets the left head panel width
	 * @return default value is 28
	 */
	public int getLeftheadwidth() {
		return _leftheadWidth;
	}
	
	/**
	 * rename setLeftheadwidth -> setLeftheadWidth
	 */
	/**
	 * Sets the left head panel width, must large then 0.
	 * @param leftWidth leaf header width
	 */
	public void setLeftheadwidth(int leftWidth) {
		if (_leftheadWidth != leftWidth) {
			_leftheadWidth = leftWidth;
			invalidate();
			/**
			 * leftw -> leftPanelWidth, not implement yet
			 */
			// smartUpdate("z.leftw", leftWidth);
			//smartUpdate("leftPanelWidth", leftWidth);
		}
	}

	/**
	 * Gets the top head panel height
	 * @return default value is 20
	 */
	public int getTopheadheight() {
		return _topheadHeight;
	}

	/**
	 * Sets the top head panel height, must large then 0.
	 * @param topHeight top header height
	 */
	public void setTopheadheight(int topHeight) {
		if (_topheadHeight != topHeight) {
			_topheadHeight = topHeight;
			invalidate();
			/**
			 * toph ->topPanelHeight -> , not implement yet
			 */
			//smartUpdate("topPanelHeight", topHeight);
		}
	}
	
	/**
	 * Sets whether show formula bar or not
	 * @param showFormulabar true if want to show formula bar
	 */
	public void setShowFormulabar(boolean showFormulabar) {
		if (_showFormulabar != showFormulabar) {
			_showFormulabar = showFormulabar;
			smartUpdate("showFormulabar", _showFormulabar);
		}
	}
	
	/**
	 * Returns whether show formula bar
	 */
	public boolean isShowFormulabar() {
		return _showFormulabar;
	}
	
	private Map convertAutoFilterToJSON(AutoFilter af) {
		if (af != null) {
			final CellRangeAddress addr = af.getRangeAddress();
			if (addr == null) {
				return null;
			}
			final Map addrmap = new HashMap();
			final int left = addr.getFirstColumn();
			final int right = addr.getLastColumn();
			final int top = addr.getFirstRow();
			final Worksheet sheet = this.getSelectedSheet();
			addrmap.put("left", left);
			addrmap.put("top", top);
			addrmap.put("right", right);
			addrmap.put("bottom", addr.getLastRow());
			
			int offcol = -1;
			final List<FilterColumn> fcs = af.getFilterColumns();
			final List<Map> fcsary = fcs != null ? new ArrayList<Map>(fcs.size()) : null;
			if (fcsary != null) {
				for(int col = left; col <= right; ++col) {
					final FilterColumn fc = af.getFilterColumn(col - left);
					final List<String> filters = fc != null ? fc.getFilters() : null;
					final boolean on = fc != null ? fc.isOn() : true;
					Map fcmap = new HashMap();
					fcmap.put("col", Integer.valueOf(col));
					fcmap.put("filter", filters);
					fcmap.put("on", on);
					int field = col - left + 1;
					if (offcol >= 0 && on) { //pre button not shown and I am shown, the field number might be different!
						field = offcol - left + 1;
					}
					fcmap.put("field", field);
					if (!on) {
						if (offcol < 0) { //first button off column
							offcol = col;
						}
					} else {
						offcol = -1;
					}
					fcsary.add(fcmap);
				}
			}
			
			final Map afmap = new HashMap();
			afmap.put("range", addrmap);
			afmap.put("filterColumns", fcsary);
			return afmap;
		}
		return null;
	}

	private List<Map> convertDataValidationToJSON(List<DataValidation> dvs) {
		if (dvs != null && !dvs.isEmpty()) {
			List<Map> results = new ArrayList<Map>(dvs.size());
			for(DataValidation dv : dvs) {
				results.add(convertDataValidationToJSON(dv));
			}
			return results;
		}
		return null;
	}
	
	private Map convertDataValidationToJSON(DataValidation dv) {
		final CellRangeAddressList addrList = dv.getRegions();
		final int addrCount = addrList.countRanges();
		final List<Map> addrmapary = new ArrayList<Map>(addrCount);
		for (int j = 0; j < addrCount; ++j) {
			final CellRangeAddress addr = addrList.getCellRangeAddress(j); 
			final int left = addr.getFirstColumn();
			final int right = addr.getLastColumn();
			final int top = addr.getFirstRow();
			final int bottom = addr.getLastRow();
			final Worksheet sheet = this.getSelectedSheet();
			final Map<String, Integer> addrmap = new HashMap<String, Integer>();
			addrmap.put("left", left);
			addrmap.put("top", top);
			addrmap.put("right", right);
			addrmap.put("bottom", bottom);
			addrmapary.add(addrmap);
		}
		final Map validMap = new HashMap();
		validMap.put("rangeList", addrmapary); //range list
		validMap.put("showButton", dv.getSuppressDropDownArrow()); //whether show dropdown button
		validMap.put("showPrompt", dv.getShowPromptBox()); //whether show prompt box
		validMap.put("promptTitle", dv.getPromptBoxTitle()); //the prompt box title
		validMap.put("promptText", dv.getPromptBoxText()); //the prompt box text
		String[] validationList = BookHelper.getValidationList(getSelectedSheet(), dv);
		if (validationList != null) {
			JSONArray jsonAry = new JSONArray();
			for (String v : validationList) {
				jsonAry.add(v);
			}
			validMap.put("validationList", jsonAry);
		}
		
		return validMap;
	}

	//ZSS-13: Support Open hyperlink in a separate browser tab window
	private boolean getLinkToNewTab() {
		final String linkToNewTab = Library.getProperty("org.zkoss.zss.ui.Spreadsheet.linkToNewTab", "true");
		return Boolean.valueOf(linkToNewTab);
	}
	
	protected void renderProperties(ContentRenderer renderer) throws IOException {
		super.renderProperties(renderer);
		renderer.render("showFormulabar", _showFormulabar);
		Worksheet sheet = this.getSelectedSheet();
		if (sheet == null) {
			return;
		}
		//handle link to new browser tab window; default to link to new tab
		if (!getLinkToNewTab()) {
			renderer.render("_linkToNewTab", false);
		}
		
		//handle AutoFilter
		Map afmap = convertAutoFilterToJSON(sheet.getAutoFilter());
		if (afmap != null) {
			renderer.render("_autoFilter", afmap);
		} else {
			renderer.render("_autoFilter", (String) null);
		}
		renderer.render("rowHeight", getRowheight());
		renderer.render("columnWidth", getColumnwidth());

		int th = isHidecolumnhead() ? 1 : this.getTopheadheight();
		
		if (_hideGridlines) {
			renderer.render("displayGridlines", !_hideGridlines);
		}
		if (_protectSheet)
			renderer.render("protect", _protectSheet);
		
		renderer.render("topPanelHeight", th);
		
		int lw = isHiderowhead() ? 1 : this.getLeftheadwidth();
		renderer.render("leftPanelWidth", lw);

		renderer.render("cellPadding", _cellpadding);
		String css = getDynamicMediaURI(this, _cssVersion++, "ss_" + this.getUuid(), "css");
		renderer.render("loadcss", new JavaScriptValue("zk.loadCSS('" + css + "', '" + this.getUuid() + "-sheet" + "')"));
		renderer.render("scss", css);

		renderer.render("maxRow", getMaxrowInJSON());
		renderer.render("maxColumn", getMaxcolumnsInJSON());

		renderer.render("sheetId", getSelectedSheetId());
		
		renderer.render("focusRect", getRectStr(_focusRect));
		renderer.render("selectionRect", getRectStr(_selectionRect));
		if (_highlightRect != null) {
			renderer.render("highLightRect", getRectStr(_highlightRect));
		}

		// generate customized row & column information
		HeaderPositionHelper colHelper = getColumnPositionHelper(sheet);
		HeaderPositionHelper rowHelper = getRowPositionHelper(sheet);
		renderer.render("csc", getSizeHelperStr(colHelper));
		renderer.render("csr", getSizeHelperStr(rowHelper));

		// generate merge range information

		MergeMatrixHelper mmhelper = getMergeMatrixHelper(sheet);
		Iterator iter = mmhelper.getRanges().iterator();
		StringBuffer merr = new StringBuffer();
		while (iter.hasNext()) {
			MergedRect block = (MergedRect) iter.next();
			int left = block.getLeft();
			int top = block.getTop();
			int right = block.getRight();
			int bottom = block.getBottom();
			int id = block.getId();
			merr.append(left).append(",").append(top).append(",").append(right).append(",").append(bottom).append(",").append(id);
			if (iter.hasNext()) {
				merr.append(";");
			}
		}
		/**
		 * mers -> mergeRange
		 */
		renderer.render("mergeRange", merr.toString());
		/**
		 * Add attr for UI renderer
		 */
		final SpreadsheetCtrl spreadsheetCtrl = ((SpreadsheetCtrl) this.getExtraCtrl());
		
		int colSize = SpreadsheetCtrl.DEFAULT_LOAD_COLUMN_SIZE;
		int preloadColSize = getPreloadColumnSize();
		if (preloadColSize == -1) {
			colSize = Math.min(colSize, getMaxcolumns() - 1);
		} else {
			colSize = Math.min(preloadColSize - 1, getMaxcolumns() - 1);
		}
		int rowSize = SpreadsheetCtrl.DEFAULT_LOAD_ROW_SIZE;
		int preloadRowSize = getPreloadRowSize();
		if (preloadRowSize == -1) {
			rowSize = Math.min(rowSize, getMaxrows() - 1);
		} else {
			rowSize = Math.min(preloadRowSize - 1, getMaxrows() - 1);
		}
		renderer.render("activeRange", 
			spreadsheetCtrl.getRangeAttrs(sheet, SpreadsheetCtrl.Header.BOTH, SpreadsheetCtrl.CellAttribute.ALL, 0, 0, 
					colSize , rowSize));
		
		renderer.render("preloadRowSize", preloadColSize);
		renderer.render("preloadColSize", preloadRowSize);
		
		renderer.render("columnHeadHidden", _hideColhead);
		renderer.render("rowHeadHidden", _hideRowhead);
		renderer.render("dataPanel", ((SpreadsheetCtrl) this.getExtraCtrl()).getDataPanelAttrs());

		//handle Validation, must after render("activeRange" ...)
		List<Map> dvs = convertDataValidationToJSON(sheet.getDataValidations());
		if (dvs != null) {
			renderer.render("dataValidations", dvs);
		} else {
			renderer.render("dataValidations", (String) null);
		}
	}

	/**
	 * Get Column title of given index, it returns column name depends on
	 * following condition<br/>
	 * 1.if there is a custom column title assign on this spreadsheet on index,
	 * then return custom value <br/>
	 * 2.if there is a custom column title assign on selected sheet on index,
	 * the return custom value <br/>
	 * 3.return default column String, for example index 0 become "A", index 9
	 * become "J"
	 * 
	 * @return column name
	 */
	public String getColumntitle(int index) {

		if (Utils.isTitleIndexMode(this)) {
			return Integer.toString(index);
		}
		String cname;
		if (_columnTitles != null
				&& (cname = (String) _columnTitles.get(Integer.valueOf(index))) != null) {
			return cname;
		}
		return CellReference.convertNumToColString(index);
	}

	/**
	 * Get Row title of given index, it returns row name depends on following
	 * condition<br/>
	 * 1.if there is a custom row title assign on this spreadsheet on index,
	 * then return custom value <br/>
	 * 2.if there is a custom row title assign on selected sheet on index, the
	 * return custom value <br/>
	 * 3.return default row index+1 String, for example index 0 become "1",
	 * index 9 become "10"
	 * 
	 * @param index row index
	 * @return row name
	 */
	public String getRowtitle(int index) {
		String rname = null;

		if (Utils.isTitleIndexMode(this)) {
			return Integer.toString(index);
		}

		if (_rowTitles != null && (rname = (String) _rowTitles.get(Integer.valueOf(index))) != null) {
			return rname;
		}

		return ""+(index+1);
	}

	/**
	 * Return current selection rectangle only if onCellSelection event listener is registered. 
	 * The returned value is a clone copy of current selection status. 
	 * Default Value:(0,0,0,0)
	 * @return current selection
	 */
	public Rect getSelection() {
		return (Rect) _selectionRect.cloneSelf();
	}

	/**
	 * Sets the selection rectangle. In general, if you set a selection, you must
	 * also set the focus by {@link #setCellFocus(Position)};. And, if you want
	 * to get the focus back to spreadsheet, call {@link #focus()} after set
	 * selection.
	 * 
	 * @param sel the selection rect
	 */
	public void setSelection(Rect sel) {
		if (!Objects.equals(_selectionRect, sel)) {
			final SpreadsheetVersion ver = _book.getSpreadsheetVersion();
			if (sel.getLeft() < 0 || sel.getTop() < 0
					|| sel.getRight() > ver.getLastColumnIndex()
					|| sel.getBottom() > ver.getLastRowIndex()
					|| sel.getLeft() > sel.getRight()
					|| sel.getTop() > sel.getBottom()) {
				throw new UiException("illegal selection : " + sel.toString());
			}
			_selectionRect.set(sel.getLeft(), sel.getTop(), sel.getRight(), sel.getBottom());
			JSONObj result = new JSONObj();
			result.setData("type", "move");
			result.setData("left", sel.getLeft());
			result.setData("top", sel.getTop());
			result.setData("right", sel.getRight());
			result.setData("bottom", sel.getBottom());

			// TODO use command to a avoid use call focusTo in .invalidate()
			// case;
			// smartUpdateValues("selection",new
			// Object[]{"",Utils.getId(getSelectedSheet()),result.toString()});
			/**
			 * rename: zssselection -> selection
			 */
			response("selection" + this.getUuid(), new AuSelection(this, result.toString()));
		}
	}

	/**
	 * Return current highlight rectangle. the returned value is a clone copy of
	 * current highlight status. Default Value: null
	 * 
	 * @return current highlight
	 */
	public Rect getHighlight() {
		if (_highlightRect == null)
			return null;
		return (Rect) _highlightRect.cloneSelf();
	}

	/**
	 * Sets the highlight rectangle or sets a null value to hide it.
	 * 
	 * @param highlight the highlight rect
	 */
	public void setHighlight(Rect highlight) {
		if (!Objects.equals(_highlightRect, highlight)) {
			JSONObj result = new JSONObj();
			if (highlight == null) {
				_highlightRect = null;
				result.setData("type", "hide");
			} else {
				final int left = Math.max(highlight.getLeft(), 0);
				final int right = Math.min(highlight.getRight(), this.getMaxcolumns()-1);
				final int top = Math.max(highlight.getTop(), 0);
				final int bottom = Math.min(highlight.getBottom(), this.getMaxrows()-1);
				if (left > right || top > bottom) {
					_highlightRect = null;
					result.setData("type", "hide");
				} else {
					_highlightRect = new Rect(left, top, right, bottom);
					result.setData("type", "show");
					result.setData("left", left);
					result.setData("top", top);
					result.setData("right", right);
					result.setData("bottom", bottom);
				}
			}
			response("selectionHighlight", new AuHighlight(this, result.toString()));
		}
	}
	
	/**
	 * Sets whether display the gridlines.
	 * @param show true to show the gridlines.
	 */
	private void setDisplayGridlines(boolean show) {
		if (_hideGridlines == show) {
			_hideGridlines = !show;
			smartUpdate("displayGridlines", show);
		}
	}
	
	/**
	 * Update autofilter buttons.
	 * @param af the current AutoFilter.
	 */
	private void updateAutoFilter(AutoFilter af) {
		smartUpdate("autoFilter", convertAutoFilterToJSON(af));
	}

    /**
     * Sets the sheet protection
     * @param boolean protect
     */
	private void setProtectSheet(boolean protect) {
		if (_protectSheet != protect) {
			_protectSheet = protect;
			smartUpdate("protect", protect);
		}
	}

	/**
	 * Return current cell(row,column) focus position. you can get the row by
	 * {@link Position#getRow()}, get the column by {@link Position#getColumn()}
	 * . The returned value is a copy of current focus status. Default
	 * Value:(0,0)
	 * 
	 * @return current focus
	 */
	public Position getCellFocus() {
		return new Position(_focusRect.getTop(), _focusRect.getLeft());
	}

	/**
	 * Sets the cell focus position.(this method doesn't focus the spreadsheet.)
	 * In general, if you set a cell focus, you also set the selection by
	 * {@link #setSelection(Rect)}; And if you want to get the focus back to
	 * spreadsheet, call {@link #focus()} to retrieve focus.
	 * 
	 * @param focus the cell focus position
	 */
	public void setCellFocus(Position focus) {
		if (_focusRect.getLeft() != focus.getColumn()
				|| _focusRect.getTop() != focus.getRow()) {
			if (focus.getColumn() < 0 || focus.getRow() < 0
					|| focus.getColumn() >= this.getMaxcolumns()
					|| focus.getRow() >= this.getMaxrows()) {
				throw new UiException("illegal position : " + focus.toString());
			}
			_focusRect.set(focus.getColumn(), focus.getRow(),
					focus.getColumn(), focus.getRow());
			JSONObj result = new JSONObj();
			result.setData("type", "move");
			result.setData("row", focus.getRow());
			result.setData("column", focus.getColumn());

			/**
			 * rename zsscellfocus -> cellFocus
			 */
			// smartUpdateValues("cellfocus",new Object[]{"",Utils.getId(getSelectedSheet()),result.toString()});
			response("cellFocus" + this.getUuid(), new AuCellFocus(this, result.toString()));
		}
	}

	/** VariableResolver to handle model's variable **/
	private class InnerVariableResolver implements VariableResolver, Serializable {
		private static final long serialVersionUID = 1L;

		public Object resolveVariable(String name) throws XelException {
			Page page = getPage();
			Object result = null;
			if (page != null) {
				result = page.getZScriptVariable(Spreadsheet.this, name);
			}
			if (result == null) {
				result = Spreadsheet.this.getAttributeOrFellow(name, true);
			}
			if (result == null && page != null) {
				result = page.getXelVariable(null, null, name, true);
			}

			return result;
		}
	}

	private class InnerFunctionMapper implements FunctionMapper, Serializable {
		private static final long serialVersionUID = 1L;

		public Collection getClassNames() {
			final Page page = getPage();
			if (page != null) {
				final FunctionMapper mapper = page.getFunctionMapper();
				if (mapper != null) {
					return mapper.getClassNames();
				}
			}
			return null;
		}

		public Class resolveClass(String name) throws XelException {
			final Page page = getPage();
			if (page != null) {
				final FunctionMapper mapper = page.getFunctionMapper();
				if (mapper != null) {
					return mapper.resolveClass(name);
				}
			}
			return null;
		}

		public Function resolveFunction(String prefix, String name)
				throws XelException {
			final Page page = getPage();
			if (page != null) {
				final FunctionMapper mapper = page.getFunctionMapper();
				if (mapper != null) {
					return mapper.resolveFunction(prefix, name);
				}
			}
			return null;
		}

	}

	private final static String[] FOCUS_COLORS = 
		new String[]{"#FFC000","#FFFF00","#92D050","#00B050","#00B0F0","#0070C0","#002060","#7030A0",
					"#4F81BD","#F29436","#9BBB59","#8064A2","#4BACC6","#F79646","#C00000","#FF0000",
					"#0000FF","#008000","#9900CC","#800080","#800000","#FF6600","#CC0099","#00FFFF"};
	
	/* DataListener to handle sheet data event */
	private class InnerDataListener extends EventDispatchListener implements Serializable {
		private static final long serialVersionUID = 20100330164021L;

		public InnerDataListener() {
			addEventListener(SSDataEvent.ON_SHEET_DELETE, new EventListener() {

				@Override
				public void onEvent(Event event) throws Exception {
					onSheetDelete((SSDataEvent)event);
				}
			});
			addEventListener(SSDataEvent.ON_FRIEND_FOCUS_MOVE, new EventListener() {
				@Override
				public void onEvent(Event event) throws Exception {
					onFriendFocusMove((SSDataEvent)event);
				}
			});
			addEventListener(SSDataEvent.ON_FRIEND_FOCUS_DELETE, new EventListener() {
				@Override
				public void onEvent(Event event) throws Exception {
					onFriendFocusDelete((SSDataEvent)event);
				}
			});
			addEventListener(SSDataEvent.ON_CONTENTS_CHANGE, new EventListener() {
				@Override
				public void onEvent(Event event) throws Exception {
					onContentChange((SSDataEvent)event);
				}
			});
			addEventListener(SSDataEvent.ON_RANGE_INSERT, new EventListener() {
				@Override
				public void onEvent(Event event) throws Exception {
					onRangeInsert((SSDataEvent)event);
				}
			});
			addEventListener(SSDataEvent.ON_SIZE_CHANGE, new EventListener() {
				@Override
				public void onEvent(Event event) throws Exception {
					onSizeChange((SSDataEvent)event);
				}
			});
			addEventListener(SSDataEvent.ON_BTN_CHANGE, new EventListener() {
				@Override
				public void onEvent(Event event) throws Exception {
					onBtnChange((SSDataEvent)event);
				}
			});
			addEventListener(SSDataEvent.ON_RANGE_DELETE, new EventListener() {
				@Override
				public void onEvent(Event event) throws Exception {
					onRangeDelete((SSDataEvent)event);
				}
			});
			addEventListener(SSDataEvent.ON_MERGE_CHANGE, new EventListener() {
				@Override
				public void onEvent(Event event) throws Exception {
					onMergeChange((SSDataEvent)event);
				}
			});
			addEventListener(SSDataEvent.ON_MERGE_ADD, new EventListener() {
				@Override
				public void onEvent(Event event) throws Exception {
					onMergeAdd((SSDataEvent)event);
				}
			});
			addEventListener(SSDataEvent.ON_MERGE_DELETE, new EventListener() {
				@Override
				public void onEvent(Event event) throws Exception {
					onMergeDelete((SSDataEvent)event);
				}
			});
			addEventListener(SSDataEvent.ON_DISPLAY_GRIDLINES, new EventListener() {
				@Override
				public void onEvent(Event event) throws Exception {
					onDisplayGridlines((SSDataEvent)event);
				}
			});
			addEventListener(SSDataEvent.ON_PROTECT_SHEET, new EventListener() {
				@Override
				public void onEvent(Event event) throws Exception {
					onProtectSheet((SSDataEvent)event);
				}
			});
			addEventListener(SSDataEvent.ON_CHART_ADD, new EventListener() {
				@Override
				public void onEvent(Event event) throws Exception {
					onChartAdd((SSDataEvent)event);
				}
			});
			addEventListener(SSDataEvent.ON_CHART_DELETE, new EventListener() {
				@Override
				public void onEvent(Event event) throws Exception {
					onChartDelete((SSDataEvent)event);
				}
			});
			addEventListener(SSDataEvent.ON_CHART_UPDATE, new EventListener() {
				@Override
				public void onEvent(Event event) throws Exception {
					onChartUpdate((SSDataEvent)event);
				}
			});
			addEventListener(SSDataEvent.ON_PICTURE_ADD, new EventListener() {
				@Override
				public void onEvent(Event event) throws Exception {
					onPictureAdd((SSDataEvent)event);
				}
			});
			addEventListener(SSDataEvent.ON_PICTURE_DELETE, new EventListener() {
				@Override
				public void onEvent(Event event) throws Exception {
					onPictureDelete((SSDataEvent)event);
				}
			});
			addEventListener(SSDataEvent.ON_PICTURE_UPDATE, new EventListener() {
				@Override
				public void onEvent(Event event) throws Exception {
					onPictureUpdate((SSDataEvent)event);
				}
			});
			addEventListener(SSDataEvent.ON_WIDGET_CHANGE, new EventListener() {
				@Override
				public void onEvent(Event event) throws Exception {
					onWidgetChange((SSDataEvent)event);
				}
			});
		}
		private void onSheetDelete(SSDataEvent event) {
			final Object[] payload = (Object[]) event.getPayload(); 
			final String delSheetName = (String) payload[0]; //deleted sheet name
			final String newSheetName= (String) payload[1]; //new selected sheet name
			org.zkoss.zk.ui.event.Events.postEvent(new SheetDeleteEvent(Events.ON_SHEET_DELETE, Spreadsheet.this, delSheetName, newSheetName));
		}
		
		private int _colorIndex = 0;
		private Worksheet getSheet(Ref rng) {
			return Utils.getSheetByRefSheet(_book, rng.getOwnerSheet()); 
		}
		private String nextFocusColor() {
			return FOCUS_COLORS[_colorIndex++ % FOCUS_COLORS.length]; 
		}
		private void onFriendFocusMove(SSDataEvent event) {
			final Ref rng = event.getRef();
			if (getSheet(rng).equals(_selectedSheet)) { //same sheet
				final Focus focus = (Focus) event.getPayload(); //other's spreadsheet's focus
				final String id = focus.id;
				if (!id.equals(_focus.id)) {
					final Focus ofocus = _focuses.get(id);
					moveEditorFocus(id, focus.name, ofocus != null ? ofocus.color : nextFocusColor(), focus.row, focus.col);
				}
			}
		}
		private void onFriendFocusDelete(SSDataEvent event) {
			final Ref rng = event.getRef();
			if (BookHelper.getSheet(_book, rng.getOwnerSheet()).equals(_selectedSheet)) { //same sheet
				final Focus focus = (Focus) event.getPayload(); //other's spreadsheet's focus
				removeEditorFocus(focus.id);
			}
		}
		private void onChartAdd(SSDataEvent event) {
			final Ref rng = event.getRef();
			final Worksheet sheet = getSheet(rng);
			final Object payload = event.getPayload();
			addChartWidget(sheet, (ZssChartX) payload);
		}
		private void onChartDelete(SSDataEvent event) {
			final Ref rng = event.getRef();
			final Worksheet sheet = getSheet(rng);
			final Object payload = event.getPayload();
			deleteChartWidget(sheet, (Chart) payload);
		}
		private void onChartUpdate(SSDataEvent event) {
			final Ref rng = event.getRef();
			final Worksheet sheet = getSheet(rng);
			final Object payload = event.getPayload();
			updateChartWidget(sheet, (Chart) payload);
		}
		private void onPictureAdd(SSDataEvent event) {
			final Ref rng = event.getRef();
			final Worksheet sheet = getSheet(rng);
			final Object payload = event.getPayload();
			addPictureWidget(sheet, (Picture) payload);
		}
		private void onPictureDelete(SSDataEvent event) {
			final Ref rng = event.getRef();
			final Worksheet sheet = getSheet(rng);
			final Object payload = event.getPayload();
			deletePictureWidget(sheet, (Picture) payload);
		}
		private void onPictureUpdate(SSDataEvent event) {
			final Ref rng = event.getRef();
			final Worksheet sheet = getSheet(rng);
			final Object payload = event.getPayload();
			updatePictureWidget(sheet, (Picture) payload);
		}
		private void onWidgetChange(SSDataEvent event) {
			final Ref rng = event.getRef();
			final Worksheet sheet = getSheet(rng);
			final int left = rng.getLeftCol();
			final int top = rng.getTopRow();
			final int right = rng.getRightCol();
			final int bottom = rng.getBottomRow();
			updateWidget(sheet, left, top, right, bottom);
		}
		private void onContentChange(SSDataEvent event) {
			final Ref rng = event.getRef();
			final Worksheet sheet = getSheet(rng);
			final int left = rng.getLeftCol();
			final int top = rng.getTopRow();
			final int right = rng.getRightCol();
			int bottom = rng.getBottomRow();
			updateWidget(sheet, left, top, right, bottom);
			updateCell(sheet, left, top, right, bottom);
			final int lastrow = sheet.getLastRowNum();
			if (bottom > lastrow) {
				bottom = lastrow;
			}
			org.zkoss.zk.ui.event.Events.postEvent(new CellSelectionEvent(Events.ON_CELL_CHANGE, Spreadsheet.this, sheet, CellSelectionEvent.SELECT_CELLS, left, top, right,  bottom));
		}
		private void onRangeInsert(SSDataEvent event) {
			final Ref rng = event.getRef();
			final Worksheet sheet = getSheet(rng);
			if (!getSelectedSheet().equals(sheet))
				return;
			_updateCellId.next();
			if (rng.isWholeColumn()) {
				final int left = rng.getLeftCol();
				final int right = rng.getRightCol();
				((ExtraCtrl) getExtraCtrl()).insertColumns(sheet, left, right - left + 1);
			} else if (rng.isWholeRow()) {
				final int top = rng.getTopRow();
				final int bottom = rng.getBottomRow();
				((ExtraCtrl) getExtraCtrl()).insertRows(sheet, top, bottom - top + 1);
			}
		}
		private void onRangeDelete(SSDataEvent event) {
			final Ref rng = event.getRef();
			final Worksheet sheet = getSheet(rng);
			if (!getSelectedSheet().equals(sheet))
				return;
			_updateCellId.next();
			if (rng.isWholeColumn()) {
				final int left = rng.getLeftCol();
				final int right = rng.getRightCol();
				((ExtraCtrl) getExtraCtrl()).removeColumns(sheet, left,
						right - left + 1);
			} else if (rng.isWholeRow()) {
				final int top = rng.getTopRow();
				final int bottom = rng.getBottomRow();
				((ExtraCtrl) getExtraCtrl()).removeRows(sheet, top, bottom - top + 1);
			}
		}
		private void onMergeChange(SSDataEvent event) {
			final Ref rng = event.getRef();
			final Ref orng = event.getOriginalRef();
			final Worksheet sheet = getSheet(orng);
			if (!getSelectedSheet().equals(sheet))
				return;
			((ExtraCtrl) getExtraCtrl()).updateMergeCell(sheet, 
					rng.getLeftCol(), rng.getTopRow(), rng.getRightCol(), rng.getBottomRow(),
					orng.getLeftCol(), orng.getTopRow(), orng.getRightCol(), orng.getBottomRow());
		}
		private void onMergeAdd(SSDataEvent event) {
			final Ref rng = event.getRef();
			final Worksheet sheet = getSheet(rng);
			if (!getSelectedSheet().equals(sheet))
				return;
			((ExtraCtrl) getExtraCtrl()).addMergeCell(sheet, 
					rng.getLeftCol(), rng.getTopRow(), rng.getRightCol(), rng.getBottomRow());

		}
		private void onMergeDelete(SSDataEvent event) {
			final Ref orng = event.getRef();
			final Worksheet sheet = getSheet(orng);
			if (!getSelectedSheet().equals(sheet))
				return;
			((ExtraCtrl) getExtraCtrl()).deleteMergeCell(sheet, orng.getLeftCol(),
					orng.getTopRow(), orng.getRightCol(), orng.getBottomRow());
		}
		private void onSizeChange(SSDataEvent event) {
			//TODO shall pass the range over to the client side and let client side do it; rather than iterate each column and send multiple command
			final Ref rng = event.getRef();
			final Worksheet sheet = getSheet(rng);
			if (!getSelectedSheet().equals(sheet))
				return;
			if (rng.isWholeColumn()) {
				final int left = rng.getLeftCol();
				final int right = rng.getRightCol();
				for (int c = left; c <= right; ++c) {
					updateColWidth(sheet, c);
				}
				final Rect rect = ((SpreadsheetCtrl) getExtraCtrl()).getVisibleRect();
				syncFriendFocusesPosition(left, rect.getTop(), rect.getRight(), rect.getBottom());
			} else if (rng.isWholeRow()) {
				final int top = rng.getTopRow();
				final int bottom = rng.getBottomRow();
				for (int r = top; r <= bottom; ++r) {
					updateRowHeight(sheet, r);
				}
				final Rect rect = ((SpreadsheetCtrl) getExtraCtrl()).getVisibleRect();
				syncFriendFocusesPosition(rect.getLeft(), top, rect.getRight(), rect.getBottom());
			}
		}
		private void onBtnChange(SSDataEvent event) {
			final Ref rng = event.getRef();
			final Worksheet sheet = getSheet(rng);
			if (!getSelectedSheet().equals(sheet))
				return;
			updateAutoFilter(sheet.getAutoFilter());
		}
		private void onDisplayGridlines(SSDataEvent event) {
			final Ref rng = event.getRef();
			final Worksheet sheet = getSheet(rng);
			if (!getSelectedSheet().equals(sheet))
				return;
			setDisplayGridlines(event.isShow());
		}
		private void onProtectSheet(SSDataEvent event) {
			final Ref rng = event.getRef();
			final Worksheet sheet = getSheet(rng);
			if (!getSelectedSheet().equals(sheet))
				return;
			setProtectSheet(event.getProtect());
		}
	}
	
	private void updateColWidth(Worksheet sheet, int col) {
		final int width = Utils.getColumnWidthInPx(sheet, col);
		final boolean newHidden = sheet.isColumnHidden(col);
		HeaderPositionHelper posHelper = getColumnPositionHelper(sheet);
		HeaderPositionInfo info = posHelper.getInfo(col);
		if ((info == null && (width != posHelper.getDefaultSize() || newHidden)) || (info != null && (info.size != width || info.hidden != newHidden))) {
			int id = info == null ? _custColId.next() : info.id;
			posHelper.setInfoValues(col, width, id, newHidden);
			((ExtraCtrl) getExtraCtrl()).setColumnWidth(sheet, col, width, id, newHidden);
		}
	}

	private void updateRowHeight(Worksheet sheet, int row) {
		final int height = Utils.getRowHeightInPx(sheet, row);
		final Row rowobj = sheet.getRow(row);
		final boolean newHidden = rowobj == null ? false : rowobj.getZeroHeight();
		HeaderPositionHelper posHelper = getRowPositionHelper(sheet);
		HeaderPositionInfo info = posHelper.getInfo(row);
		if ((info == null && (height != posHelper.getDefaultSize() || newHidden)) || (info != null && (info.size != height || info.hidden != newHidden))) {
			int id = info == null ? _custRowId.next() : info.id;
			posHelper.setInfoValues(row, height, id, newHidden);
			((ExtraCtrl) getExtraCtrl()).setRowHeight(sheet, row, height, id, newHidden);
		}
	}
	
	public MergeMatrixHelper getMergeMatrixHelper(Worksheet sheet) {
		if (sheet != getSelectedSheet())
			throw new UiException("not current selected sheet ");
		MergeMatrixHelper mmhelper = (MergeMatrixHelper) getAttribute(MERGE_MATRIX_KEY);
		int fzr = getRowfreeze();
		int fzc = getColumnfreeze();
		if (mmhelper == null) {
			final int sz = sheet.getNumMergedRegions();
			final List<int[]> mergeRanges = new ArrayList<int[]>(sz);
			for(int j = sz - 1; j >= 0; --j) {
				final CellRangeAddress addr = sheet.getMergedRegion(j);
				mergeRanges.add(new int[] {addr.getFirstColumn(), addr.getFirstRow(), addr.getLastColumn(), addr.getLastRow()});
			}
			mmhelper = new MergeMatrixHelper(mergeRanges, fzr, fzc);
			setAttribute(MERGE_MATRIX_KEY, mmhelper);
		} else {
			mmhelper.update(fzr, fzc);
		}
		return mmhelper;
	}
	
	private HeaderPositionHelper getRowPositionHelper(Worksheet sheet) {
		final HeaderPositionHelper[] helper = getPositionHelpers(sheet);
		return helper != null ? helper[0] : null;
	}
	
	private HeaderPositionHelper getColumnPositionHelper(Worksheet sheet) {
		final HeaderPositionHelper[] helper = getPositionHelpers(sheet); 
		return helper != null ? helper[1] : null;
	}
	
	//[0] row position, [1] column position
	private HeaderPositionHelper[] getPositionHelpers(Worksheet sheet) {
		if (sheet == null) {
			return null;
		}
		if (sheet != getSelectedSheet())
			throw new UiException("not current selected sheet: "+sheet.getSheetName());
		HeaderPositionHelper helper = (HeaderPositionHelper) getAttribute(ROW_SIZE_HELPER_KEY);

		int maxcol = 0;
		if (helper == null) {
			int defaultSize = this.getRowheight();
			
			List<HeaderPositionInfo> infos = new ArrayList<HeaderPositionInfo>();
			
			for(Row row: sheet) {
				final boolean hidden = row.getZeroHeight();
				final int height = Utils.getRowHeightInPx(sheet, row);
				if (height != defaultSize || hidden) { //special height or hidden
					infos.add(new HeaderPositionInfo(row.getRowNum(), height, _custRowId.next(), hidden));
				}
				final int colnum = row.getLastCellNum() - 1;
				if (colnum > maxcol) {
					maxcol = colnum;
				}
			}
			
			helper = new HeaderPositionHelper(defaultSize, infos);

			setAttribute(ROW_SIZE_HELPER_KEY, helper);
		}
		return new HeaderPositionHelper[] {helper, myGetColumnPositionHelper(sheet, maxcol)};
	}

	/**
	 * Update cell data/format of selected sheet to client side, you must assign
	 * a block from left-top to right-bottom.
	 * 
	 * @param left left of block
	 * @param top top of block
	 * @param right right of block
	 * @param bottom bottom of block
	 */
	/* package */void updateCell(int left, int top, int right, int bottom) {
		updateCell(getSelectedSheet(), left, top, right, bottom);
	}

	private void updateWidget(Worksheet sheet, int left, int top, int right, int bottom) {
		if (this.isInvalidated())
			return;// since it is invalidate, we don't need to do anymore
		//update widgets per the content change of the range.
		getWidgetHandler().updateWidgets(sheet, left, top, right, bottom);
	}
	
	private void updateCell(Worksheet sheet, int left, int top, int right, int bottom) {
		if (this.isInvalidated())
			return;// since it is invalidate, we don't need to do anymore

		String sheetId = Utils.getSheetUuid(sheet);
		if (!sheetId.equals(this.getSelectedSheetId()))
			return;
		left = left > 0 ? left - 1 : 0;// for border, when update a range, we
		// should also update the left - 1, top - 1 part
		top = top > 0 ? top - 1 : 0;

		final int loadLeft = _loadedRect.getLeft();
		final int loadTop = _loadedRect.getTop();
		final int loadRight = _loadedRect.getRight();
		final int loadBottom = _loadedRect.getBottom();
		
		final int frRow = getRowfreeze();
		final int frCol = getColumnfreeze();
		
		final int frTop = top <= frRow ? top : -1;
		final int frBottom = frRow;
		final int frLeft = left <= frCol ? left : -1;
		final int frRight = frCol;
		
		if (loadLeft > left) {
			left = loadLeft;
		}
		if (loadRight < right) {
			right = loadRight;
		}
		if (loadTop > top) {
			top = loadTop;
		}
		if (loadBottom < bottom) {
			bottom = loadBottom; 
		}
		
		//TODO: update freeze range
		//row freeze part
		if (frTop >= 0 && frTop <= frBottom && left >= 0 && left <= right) { 
			responseUpdateCell(sheet, sheetId, left, frTop, right, frBottom);
		}
		//column freeze part
		if (frLeft >= 0 && frLeft <= frRight && top >= 0 && top <= bottom) {
			responseUpdateCell(sheet, sheetId, frLeft, top, frRight, bottom);
		}
		//loaded rect
		if (top >= 0 && top <= bottom && left >= 0 && left <= right) {
			responseUpdateCell(sheet, sheetId, left, top, right, bottom);
		}
	}
	
	/**
	 * Internal Use Only
	 */
	public void updateRange(Worksheet sheet, String sheetId, int left, int top, int right, int bottom) {
		SpreadsheetCtrl ctrl = (SpreadsheetCtrl) getExtraCtrl();
		String ret = ctrl.getRangeAttrs(sheet, SpreadsheetCtrl.Header.NONE, SpreadsheetCtrl.CellAttribute.ALL, left, top, right, bottom).toJSONString();
		response(bottom + "_" + right + "_" + _updateRangeId.next(), new AuUpdateData(this, AuUpdateData.UPDATE_RANGE_FUNCTION, "", sheetId, ret));
	}
	
	public void escapeAndUpdateText(Cell cell, String text) {
		final CellStyle style = (cell == null) ? null : cell.getCellStyle();
		final boolean wrap = style != null && style.getWrapText();
		text = Utils.escapeCellText(text, wrap, true);
		updateText(cell, text);
	}
	
	/**
	 * Internal Use Only
	 */
	public void updateText(Cell cell, String text) {
		if (cell == null)
			return;
		final int row = cell.getRowIndex();
		final int col = cell.getColumnIndex();
		// update cell only in block or in freeze panels
		if (row > _loadedRect.getBottom()
				|| (row < _loadedRect.getTop() && row > getRowfreeze()))
			return;
		if (col > _loadedRect.getRight()
				|| (col < _loadedRect.getLeft() && row > getColumnfreeze()))
			return;

		// if(cell==null) continue;

		final JSONObj result = new JSONObj();
		result.setData("r", row);
		result.setData("c", col);
		result.setData("type", "udtext");
		result.setData("val", text);
		final String sheetId = Utils.getSheetUuid((Worksheet) cell.getSheet());
		response(row + "_" + col + "_" + _updateCellId.next(), new AuDataUpdate(this, "", sheetId, result));
	}
	
	private void responseUpdateCell(Worksheet sheet, String sheetId, int left, int top, int right, int bottom) {
		SpreadsheetCtrl spreadsheetCtrl = ((SpreadsheetCtrl) this.getExtraCtrl());
		JSONObject result = spreadsheetCtrl.getRangeAttrs(sheet, SpreadsheetCtrl.Header.NONE, CellAttribute.ALL, left, top, right, bottom);
		result.put("type", "udcell");
		response(top + "_" + left + "_" + _updateCellId.next(), 
				new AuDataUpdate(this, "", sheetId, result));
	}
	
	private HeaderPositionHelper myGetColumnPositionHelper(Worksheet sheet, int maxcol) {
		if (sheet != getSelectedSheet())
			throw new UiException("not current selected sheet ");
		HeaderPositionHelper helper = (HeaderPositionHelper) getAttribute(COLUMN_SIZE_HELPER_KEY);

		if (helper == null) {
			final int defaultColSize = sheet.getDefaultColumnWidth();
			final int defaultColSize256 = defaultColSize * 256; 
			final int charWidth = getDefaultCharWidth();
			final int defaultColSizeInPx = Utils.defaultColumnWidthToPx(defaultColSize, charWidth);
			List<HeaderPositionInfo> infos = new ArrayList<HeaderPositionInfo>();
			for(int j=0; j <= maxcol; ++j) {
				final boolean hidden = sheet.isColumnHidden(j); //whether this column is hidden
				final int fileColumnWidth = sheet.getColumnWidth(j); //file column width
				if (fileColumnWidth != defaultColSize256 || hidden) {
					final int colwidth = fileColumnWidth != defaultColSize256 ? 
							Utils.fileChar256ToPx(fileColumnWidth, charWidth) : defaultColSizeInPx; 
					infos.add(new HeaderPositionInfo(j, colwidth, _custColId.next(), hidden));
				}
			}
			helper = new HeaderPositionHelper(defaultColSizeInPx, infos);
			setAttribute(COLUMN_SIZE_HELPER_KEY, helper);
		}
		return helper;
	}
	
	
	@Override
	public Object getExtraCtrl() {
		return newExtraCtrl();
	}

	/**
	 * Return a extra controller. only spreadsheet developer need to call this
	 * method.
	 */
	@Override
	protected Object newExtraCtrl() {
		return new ExtraCtrl();
	}

	private class ExtraCtrl implements SpreadsheetCtrl, SpreadsheetInCtrl,
			SpreadsheetOutCtrl, DynamicMedia {

		public Media getMedia(String pathInfo) {
			return new AMedia("css", "css", "text/css;charset=UTF-8",
					getSheetDefaultRules());
		}

		public void setColumnSize(String sheetId, int column, int newsize, int id, boolean hidden) {
			Worksheet sheet;
			if (getSelectedSheetId().equals(sheetId)) {
				sheet = getSelectedSheet();
			} else {
				sheet = Utils.getSheetByUuid(_book, sheetId);
			}
			// update helper size first before sheet.setColumnWidth, or it will fire a SSDataEvent
			HeaderPositionHelper helper = Spreadsheet.this.getColumnPositionHelper(sheet);
			helper.setInfoValues(column, newsize, id, hidden);

			Ranges.range(sheet, 0, column).setColumnWidth(Utils.pxToFileChar256(newsize, ((Book)sheet.getWorkbook()).getDefaultCharWidth()));
			//sheet.setColumnWidth(column, Utils.pxToFileChar256(newsize, ((Book)sheet.getWorkbook()).getDefaultCharWidth()));
		}

		public void setRowSize(String sheetId, int rownum, int newsize, int id, boolean hidden) {
			Worksheet sheet;
			if (getSelectedSheetId().equals(sheetId)) {
				sheet = getSelectedSheet();
			} else {
				sheet = Utils.getSheetByUuid(_book, sheetId);
			}
			//Row row = sheet.getRow(rownum);
			//if (row == null) {
			//	row = sheet.createRow(rownum);
			//}
			Ranges.range(sheet, rownum, 0).setRowHeight(Utils.pxToPoint(newsize));
			//row.setHeight((short)Utils.pxToTwip(newsize));
			HeaderPositionHelper helper = Spreadsheet.this.getRowPositionHelper(sheet);
			helper.setInfoValues(rownum, newsize, id, hidden);
		}

		public HeaderPositionHelper getColumnPositionHelper(String sheetId) {
			Worksheet sheet;
			if (getSelectedSheetId().equals(sheetId)) {
				sheet = getSelectedSheet();
			} else {
				sheet = Utils.getSheetByUuid(_book, sheetId);
			}
			HeaderPositionHelper helper = Spreadsheet.this
					.getColumnPositionHelper(sheet);
			return helper;
		}

		public HeaderPositionHelper getRowPositionHelper(String sheetId) {
			Worksheet sheet;
			if (getSelectedSheetId().equals(sheetId)) {
				sheet = getSelectedSheet();
			} else {
				sheet = Utils.getSheetByUuid(_book, sheetId);
			}
			HeaderPositionHelper helper = Spreadsheet.this
					.getRowPositionHelper(sheet);
			return helper;
		}

		public MergeMatrixHelper getMergeMatrixHelper(Worksheet sheet) {
			return Spreadsheet.this.getMergeMatrixHelper(sheet);
		}

		public Rect getSelectionRect() {
			return (Rect) _selectionRect.cloneSelf();
		}

		public Rect getFocusRect() {
			return (Rect) _focusRect.cloneSelf();
		}

		public void setSelectionRect(int left, int top, int right, int bottom) {
			_selectionRect.set(left, top, right, bottom);
		}

		public void setFocusRect(int left, int top, int right, int bottom) {
			_focusRect.set(left, top, right, bottom);
		}

		public Rect getLoadedRect() {
			return (Rect) _loadedRect.cloneSelf();
		}

		public void setLoadedRect(int left, int top, int right, int bottom) {
			_loadedRect.set(left, top, right, bottom);
			getWidgetHandler().onLoadOnDemand(getSelectedSheet(), left, top, right, bottom);
		}
		
		public void setVisibleRect(int left, int top, int right, int bottom) {
			_visibleRect.set(left, top, right, bottom);
			getWidgetHandler().onLoadOnDemand(getSelectedSheet(), left, top, right, bottom);
		}

		public Rect getVisibleRect() {
			return (Rect) _visibleRect.cloneSelf();
		}

		public boolean addWidget(Widget widget) {
			return Spreadsheet.this.addWidget(widget);
		}

		public boolean removeWidget(Widget widget) {
			return Spreadsheet.this.removeWidget(widget);
		}
		
		public WidgetHandler getWidgetHandler() {
			return Spreadsheet.this.getWidgetHandler();
		}
		
		public JSONObject getRowHeaderAttrs(Worksheet sheet, int rowStart, int rowEnd) {
			return getHeaderAttrs(sheet, true, rowStart, rowEnd);
		}
		
		public JSONObject getColumnHeaderAttrs(Worksheet sheet, int colStart, int colEnd) {
			return getHeaderAttrs(sheet, false, colStart, colEnd);
		}
		
		/**
		 * Header attributes
		 * 
		 * <ul>
		 * 	<li>t: header type</li>
		 *  <li>s: index start</li>
		 *  <li>e: index end</li>
		 *  <li>hs: headers, a JSONArray object</li>
		 * </ul>
		 * 
		 * @param isRow
		 * @param start
		 * @param end
		 * @return
		 */
		private JSONObject getHeaderAttrs(Worksheet sheet, boolean isRow, int start, int end) {
			JSONObject attrs = new JSONObject();
			attrs.put("s", start);
			attrs.put("e", end);
			
			JSONArray headers = new JSONArray();
			attrs.put("hs", headers);
			if (isRow) {
				attrs.put("t", "r");
				for (int row = start; row <= end; row++) {
					headers.add(getRowHeaderAttrs(sheet, row));
				}
			} else { //column header
				attrs.put("ht", "c");
				for (int col = start; col <= end; col++) {
					headers.add(getColumnHeaderAttrs(sheet, col));
				}
			}
			return attrs;
		}
		
		/**
		 * Column header attributes
		 * 
		 * <ul>
		 * 	<li>i: column index</li>
		 *  <li>t: title</li>
		 *  <li>p: position info id</li>
		 * </ul>
		 * 
		 * Ignore attribute if it's default
		 * Default attributes
		 * <ul>
		 * 	<li>hidden: false</li>
		 * </ul>
		 * 
		 * @return
		 */
		private JSONObject getColumnHeaderAttrs(Worksheet sheet, int col) {
			JSONObject attrs = new JSONObject();
			attrs.put("i", col);
			attrs.put("t", Spreadsheet.this.getColumntitle(col));

			HeaderPositionHelper colHelper = Spreadsheet.this.getColumnPositionHelper(sheet);
			HeaderPositionInfo info = colHelper.getInfo(col);
			if (info != null) {
				attrs.put("p", info.id);
//				if (info.size != defaultSize) {
//					attrs.put("s", info.size);
//				}
//				if (info.hidden) {
//					attrs.put("h", 1); //1 stand for true;
//				}
			}
			return attrs;
		}
		
		/**
		 * Row header attributes
		 * 
		 * <ul>
		 * 	<li>i: row index</li>
		 *  <li>t: title</li>
		 *  <li>p: position info id</li>
		 * </ul>
		 * 
		 * Ignore attribute if it's default
		 * Default attributes
		 * <ul>
		 * 	<li>hidden: false</li>
		 * </ul>
		 * 
		 * @return
		 */
		private JSONObject getRowHeaderAttrs(Worksheet sheet, int row) {
			JSONObject attrs = new JSONObject();
			attrs.put("i", row);
			attrs.put("t", Spreadsheet.this.getRowtitle(row));

			HeaderPositionHelper rowHelper = Spreadsheet.this.getRowPositionHelper(sheet);
			HeaderPositionInfo info = rowHelper.getInfo(row);
			if (info != null) {
				attrs.put("p", info.id);
//				if (info.hidden)
//					attrs.put("h", 1);
			}
			return attrs;
		}

		/**
		 * Range attributes
		 * 
		 * <ul>
		 * 	<li>l: range top</li>
		 *  <li>t: range top</li>
		 *  <li>r: range right</li>
		 *  <li>b: range bottom</li>
		 *  <li>at: range update Attribute Type</li>
		 *  <li>rs: rows, a JSONArray object</li>
		 *  <li>cs: cells, a JSONArray object</li>
		 * 	<li>s: strings, a JSONArray object</li>
		 *  <li>st: styles, a JSONArray object</li>
		 *  <li>m: merge attributes</li>
		 *  <li>rhs: row headers, a JSONArray object</li>
		 *  <li>chs: column headers, a JSONArray object</li>
		 * </ul>
		 * 
		 * @param left
		 * @param top
		 * @param right
		 * @param bottom
		 * @param containsHeader
		 * @return
		 */
		public JSONObject getRangeAttrs(Worksheet sheet, Header containsHeader, CellAttribute type, int left, int top, int right, int bottom) {
			JSONObject attrs = new JSONObject();
			
			attrs.put("l", left);
			attrs.put("t", top);
			attrs.put("r", right);
			attrs.put("b", bottom);
			attrs.put("at", type);
			
			JSONArray rows = new JSONArray();
			attrs.put("rs", rows);
			
			StringAggregation styleAggregation = new StringAggregation();
			StringAggregation textAggregation = new StringAggregation();
			MergeAggregation mergeAggregation = new MergeAggregation(getMergeMatrixHelper(sheet));
			for (int row = top; row <= bottom; row++) {
				JSONObject r = getRowAttrs(row);
				rows.add(r);
				
				JSONArray cells = new JSONArray();
				r.put("cs", cells);
				for (int col = left; col <= right; col++) {
					cells.add(getCellAttr(sheet, type, row, col, styleAggregation, textAggregation, mergeAggregation));
				}
			}
			
			attrs.put("s", textAggregation.getJSONArray());
			attrs.put("st", styleAggregation.getJSONArray());
			attrs.put("m", mergeAggregation.getJSONObject());
			
			boolean addRowColumnHeader = containsHeader == Header.BOTH;
			boolean addRowHeader = addRowColumnHeader || containsHeader == Header.ROW;
			boolean addColumnHeader = addRowColumnHeader || containsHeader == Header.COLUMN;
			
			if (addRowHeader)
				attrs.put("rhs", getRowHeaderAttrs(sheet, top, bottom));
			if (addColumnHeader)
				attrs.put("chs", getColumnHeaderAttrs(sheet, left, right));
			
			return attrs;
		}
		
		/**
		 * Row attributes
		 * <ul>
		 * 	<li>r: row number</li>
		 *  <li>h: height index</li>
		 *  <li>hd: hidden</li>
		 * </ul>
		 * 
		 * Ignore if attribute is default
		 * <ul>
		 * 	<li>hidden: default is false</li>
		 * </ul>
		 */
		public JSONObject getRowAttrs(int row) {
			Worksheet sheet = getSelectedSheet();
			HeaderPositionHelper helper = Spreadsheet.this.getRowPositionHelper(sheet);
			JSONObject attrs = new JSONObject();
			//row num
			attrs.put("r", row);
			
			HeaderPositionInfo info = helper.getInfo(row);
			if (info != null) {
				attrs.put("h", info.id);
				if (info.hidden) {
					attrs.put("hd", "t"); //t stand for true
				}
			}
			return attrs;
		}

		/**
		 * 
		 * Cell attributes
		 * 
		 * <ul>
		 * 	<li>r: row number</li>
		 *  <li>c: column number</li>
		 *  <li>w: width index</li>
		 *  <li>h: height index</li>
		 *  <li>t: cell html text</li>
		 *  <li>et: cell edit text</li>
		 *  <li>ft: format text</li>
		 *  <li>meft: merge cell html text, edit text and format text</li>
		 *  <li>ct: cell type</li>
		 *  <li>s: cell style</li>
		 *  <li>is: cell inner style</li>
		 *  <li>rb: cell right border</li>
		 *  <li>l: locked</>
		 *  <li>wp: wrap</li>
		 *  <li>ha: horizontal alignment</>
		 *  <li>va: vertical alignment</>
		 *  <li>mi: merge id index</li>
		 *  <li>mc: merge CSS index</li>
		 *  <li>rf: Cell Reference</li>
		 * </ul>
		 * 
		 * Ignore put attribute if it's default
		 * Default attributes
		 * <ul>
		 * 	<li>Cell type: blank</>
		 *  <li>Locked: true</>
		 *  <li>Wrap: false</li>
		 *  <li>Horizontal alignment: left</>
		 *  <li>Vertical alignment: top</>
		 * </ul>
		 */
		public JSONObject getCellAttr(Worksheet sheet, CellAttribute type, int row, int col, StringAggregation styleAggregation, StringAggregation textAggregation, MergeAggregation mergeAggregation) {
			boolean updateAll = type == CellAttribute.ALL,
				updateText = (updateAll || type == CellAttribute.TEXT),
				updateStyle = (updateAll || type == CellAttribute.STYLE),
				updateSize = (updateAll || type == CellAttribute.SIZE),
				updateMerge = (updateAll || type == CellAttribute.MERGE);
			
			Cell cell = Utils.getCell(sheet, row, col);
			JSONObject attrs = new JSONObject();
			
			//row num, cell num attr
//			if (cell != null) {
//				attrs.put("r", row);
//				attrs.put("c", col);
//			}
			CellReference cr = new CellReference(row, col, false, false);
			attrs.put("rf", cr.formatAsString());
			
			//width, height id
			if (updateSize) {
				HeaderPositionHelper colHelper = Spreadsheet.this.getColumnPositionHelper(sheet);
				HeaderPositionInfo info = colHelper.getInfo(col);
				if (info != null && info.id >= 0) {
					attrs.put("w", info.id);
				}
			}
			
			//merge attr
			if (updateMerge) {
				MergeIndex mergeIndex = mergeAggregation.add(row, col);
				if (mergeIndex != null) {
					attrs.put("mi", mergeIndex.getMergeId());
					attrs.put("mc", mergeIndex.getMergeCSSId());
				}
			}
			
			//style attr
			if (updateStyle) {
				CellFormatHelper cfh = new CellFormatHelper(sheet, row, col, getMergeMatrixHelper(sheet));
				String style = cfh.getHtmlStyle();
				if (!Strings.isEmpty(style)) {
					int idx = styleAggregation.add(style);
					attrs.put("s", idx);
				}
				String innerStyle = cfh.getInnerHtmlStyle();
				if (!Strings.isEmpty(innerStyle)) {
					int idx = styleAggregation.add(innerStyle);
					attrs.put("is", idx);
				}
				if (cfh.hasRightBorder()) {
					attrs.put("rb", 1); 
				}
			}
			
			if (cell != null) {
				int cellType = cell.getCellType();
				if (cellType != Cell.CELL_TYPE_BLANK)
					attrs.put("ct", cellType);
				
				if (updateText) {
					if (cellType != Cell.CELL_TYPE_BLANK) {
						final String cellText = getCelltext(Spreadsheet.this, row, col);
						final String editText = getEdittext(Spreadsheet.this, row, col);
						final String formatText = getCellFormatText(Spreadsheet.this, row, col);
						
						if (Objects.equals(cellText, editText) && Objects.equals(editText, formatText)) {
							attrs.put("meft", textAggregation.add(cellText));
						} else {
							attrs.put("t", textAggregation.add(cellText));
							attrs.put("et", textAggregation.add(editText));
							attrs.put("ft", textAggregation.add(formatText));
						}
					}
				}
				
				if (updateStyle) {
					CellStyle cellStyle = cell.getCellStyle();
					boolean locked = cellStyle.getLocked();
					if (!locked)
						attrs.put("l", "f"); //f stand for "false"
					
					boolean wrap = cellStyle.getWrapText();
					if (wrap)
						attrs.put("wp", 1);
					
					int horizontalAlignment = BookHelper.getRealAlignment(cell);
					switch(horizontalAlignment) {
					case CellStyle.ALIGN_CENTER:
					case CellStyle.ALIGN_CENTER_SELECTION:
						attrs.put("ha", "c");
						break;
					case CellStyle.ALIGN_RIGHT:
						attrs.put("ha", "r");
						break;
					}
					
					int verticalAlignment = cellStyle.getVerticalAlignment();
					switch(verticalAlignment) {
					case CellStyle.VERTICAL_TOP:
						attrs.put("va", "t");
						break;
					case CellStyle.VERTICAL_CENTER:
						attrs.put("va", "c");
						break;
					//case CellStyle.VERTICAL_BOTTOM: //default
					//	break;
					}
				}
			}
			return attrs;
		}

		/**
		 * API for implementation, gets data panel attributes, only spreadsheet
		 * developer need to call this method.
		 * 
		 * @return attributes
		 */
		public String getDataPanelAttrs() {
			final StringBuffer sb = new StringBuffer(64);

			HTMLs.appendAttribute(sb, "z.w", getDataPanelWidth());
			HTMLs.appendAttribute(sb, "z.h", getDataPanelHeight());
			return sb.toString();
		}

		public void insertColumns(Worksheet sheet, int col, int size) {

			if (size <= 0) {
				throw new UiException("size must > 0 : " + size);
			}
			if (col > _maxColumns) {// not in scope, do nothing,
				return;
			}
			// because of rowfreeze and colfreeze,
			// don't avoid insert behavior here, always send required data to
			// client,
			// let client handle it

			// remove merge before a new column or row
			removeAffectedMergeRange(sheet, 0, col);

			JSONObject result = new JSONObject();
			result.put("type", "column");
			result.put("col", col);
			result.put("size", size);

			//TODO: seems no need at all ???
			List extnm = new ArrayList();
			int right = size + _loadedRect.getRight();
			for (int i = col; i <= right; i++) {
				extnm.add(getColumntitle(i));
			}
			result.put("extnm", extnm);

			HeaderPositionHelper colHelper = Spreadsheet.this
					.getColumnPositionHelper(sheet);

			colHelper.shiftMeta(col, size);
			//_maxColumns += size;
			int cf = getColumnfreeze();
			if (cf >= col) {
				_colFreeze += size;
			}

			result.put("maxcol", _maxColumns);
			result.put("colfreeze", _colFreeze);

			response("insertRowColumn" + Utils.nextUpdateId(), new AuInsertRowColumn(Spreadsheet.this, "", Utils.getSheetUuid(sheet), result));

			_loadedRect.setRight(right);

			// update surround cell
			int left = col;
			right = left + size - 1;
			right = right >= _maxColumns - 1 ? _maxColumns - 1 : right;
			int top = _loadedRect.getTop();
			int bottom = _loadedRect.getBottom();
			
			log.debug("update cells when insert column " + col + ",size:" + size + ":" + left + "," + top + "," + right + "," + bottom);
			updateCell(sheet, left, top, right, bottom);
			
			//update inserted column widths
			updateColWidths(sheet, col, size); 
		}

		private void updateRowHeights(Worksheet sheet, int row, int n) {
			for(int r = 0; r < n; ++r) {
				updateRowHeight(sheet, r+row);
			}
		}
		
		private void updateColWidths(Worksheet sheet, int col, int n) {
			for(int r = 0; r < n; ++r) {
				updateColWidth(sheet, r+col);
			}
		}
		public void insertRows(Worksheet sheet, int row, int size) {
			if (size <= 0) {
				throw new UiException("size must > 0 : " + size);
			}
			if (row > _maxRows) {// not in scrope, do nothing,
				return;
			}
			// because of rowfreeze and colfreeze,
			// don't avoid insert behavior here, always send required data to
			// client,
			// let client handle it

			// remove merge before a new column or row
			removeAffectedMergeRange(sheet, 1, row);

			JSONObject result = new JSONObject();
			result.put("type", "row");
			result.put("row", row);
			result.put("size", size);

			List extnm = new ArrayList();
			int bottom = size + _loadedRect.getBottom();
			for (int i = row; i <= bottom; i++) {
				extnm.add(getRowtitle(i));
			}
			result.put("extnm", extnm);

			HeaderPositionHelper rowHelper = Spreadsheet.this.getRowPositionHelper(sheet);
			rowHelper.shiftMeta(row, size);
			//_maxRows += size;
			int rf = getRowfreeze();
			if (rf >= row) {
				_rowFreeze += size;
			}

			result.put("maxrow", _maxRows);
			result.put("rowfreeze", _rowFreeze);

			response("insertRowColumn" + Utils.nextUpdateId(), new AuInsertRowColumn(Spreadsheet.this, "", Utils.getSheetUuid(sheet), result));

			_loadedRect.setBottom(bottom);

			// update surround cell
			int top = row;
			bottom = bottom + size - 1;
			bottom = bottom >= _maxRows - 1 ? _maxRows - 1 : bottom;
			updateCell(sheet, _loadedRect.getLeft(), top, _loadedRect.getRight(), bottom);
			
			// update the inserted row height
			updateRowHeights(sheet, row, size); //update row height
		}

		public void removeColumns(Worksheet sheet, int col, int size) {
			if (size <= 0) {
				throw new UiException("size must > 0 : " + size);
			} else if (col < 0) {
				throw new UiException("column must >= 0 : " + col);
			}
			if (col >= _maxColumns) {
				return;
			}

			if (col + size > _maxColumns) {
				size = _maxColumns - col;
			}

			// because of rowfreeze and colfreeze,
			// don't avoid insert behavior here, always send required data to
			// client,
			// let client handle it

			// remove merge before a new column or row
			removeAffectedMergeRange(sheet, 0, col);

			JSONObj result = new JSONObj();
			result.setData("type", "column");
			result.setData("col", col);
			result.setData("size", size);

			List extnm = new ArrayList();

			int right = _loadedRect.getRight() - size;
			if (right < col) {
				right = col - 1;
			}
			for (int i = col; i <= right; i++) {
				extnm.add(getColumntitle(i));
			}
			result.setData("extnm", extnm);

			HeaderPositionHelper colHelper = Spreadsheet.this.getColumnPositionHelper(sheet);
			colHelper.unshiftMeta(col, size);

			//_maxColumns -= size;
			int cf = getColumnfreeze();
			if (cf > -1 && col <= cf) {
				if (col + size > cf) {
					_colFreeze = col - 1;
				} else {
					_colFreeze -= size;
				}
			}

			result.setData("maxcol", _maxColumns);
			result.setData("colfreeze", _colFreeze);

			response("removeRowColumn" + Utils.nextUpdateId(), new AuRemoveRowColumn(Spreadsheet.this, "", Utils.getSheetUuid(sheet), result));
			_loadedRect.setRight(right);

			// update surround cell
			int left = col;
			right = left;
			
			updateCell(sheet, left, _loadedRect.getTop(), right, _loadedRect.getBottom());
		}

		public void removeRows(Worksheet sheet, int row, int size) {
			if (size <= 0) {
				throw new UiException("size must > 0 : " + size);
			} else if (row < 0) {
				throw new UiException("row must >= 0 : " + row);
			}
			if (row >= _maxRows) {
				return;
			}
			if (row + size > _maxRows) {
				size = _maxRows - row;
			}

			// because of rowfreeze and colfreeze,
			// don't avoid insert behavior here, always send required data to
			// client,
			// let client handle it

			// remove merge before a new column or row
			removeAffectedMergeRange(sheet, 1, row);

			JSONObject result = new JSONObject();
			result.put("type", "row");
			result.put("row", row);
			result.put("size", size);

			List extnm = new ArrayList();

			int bottom = _loadedRect.getBottom() - size;
			if (bottom < row) {
				bottom = row - 1;
			}
			for (int i = row; i <= bottom; i++) {
				extnm.add(getRowtitle(i));
			}
			result.put("extnm", extnm);

			HeaderPositionHelper rowHelper = Spreadsheet.this.getRowPositionHelper(sheet);
			rowHelper.unshiftMeta(row, size);

//			_maxRows -= size;
			int rf = getRowfreeze();
			if (rf > -1 && row <= rf) {
				if (row + size > rf) {
					_rowFreeze = row - 1;
				} else {
					_rowFreeze -= size;
				}
			}

			result.put("maxrow", _maxRows);
			result.put("rowfreeze", _rowFreeze);

			response("removeRowColumn" + Utils.nextUpdateId(), new AuRemoveRowColumn(Spreadsheet.this, "", Utils.getSheetUuid(sheet), result));
			_loadedRect.setBottom(bottom);

			// update surround cell
			int top = row;
			bottom = top;
			
			updateCell(sheet, _loadedRect.getLeft(), top, _loadedRect.getRight(), bottom);
		}

		private void removeAffectedMergeRange(Worksheet sheet, int type, int index) {
//handled by onMergeChange... 			
/*			
			MergeMatrixHelper mmhelper = this.getMergeMatrixHelper(sheet);
			List toremove = new ArrayList();
			if (type == 0) {// column
				mmhelper.deleteAffectedMergeRangeByColumn(index, toremove);
			} else if (type == 1) {
				mmhelper.deleteAffectedMergeRangeByRow(index, toremove);
			} else {
				return;
			}
			for (Iterator iter = toremove.iterator(); iter.hasNext();) {
				MergedRect block = (MergedRect) iter.next();
				updateMergeCell0(sheet, block, "remove");
			}
*/		}

		public void updateMergeCell(Worksheet sheet, int left, int top, int right,
				int bottom, int oleft, int otop, int oright, int obottom) {
			deleteMergeCell(sheet, oleft, otop, oright, obottom);
			addMergeCell(sheet, left, top, right, bottom);
		}

		public void deleteMergeCell(Worksheet sheet, int left, int top, int right, int bottom) {
			MergeMatrixHelper mmhelper = this.getMergeMatrixHelper(sheet);
			Set torem = new HashSet();
			mmhelper.deleteMergeRange(left, top, right, bottom, torem);
			for (Iterator iter = torem.iterator(); iter.hasNext();) {
				MergedRect rect = (MergedRect) iter.next();

				updateMergeCell0(sheet, rect, "remove");
			}
			//updateCell(sheet, left > 0 ? left - 1 : 0, top > 1 ? top - 1 : 0, right + 1, bottom + 1);
			updateCell(sheet, left, top, right, bottom);
		}

		private void updateMergeCell0(Worksheet sheet, MergedRect block, String type) {
			JSONObj result = new JSONObj();
			result.setData("type", type);
			result.setData("id", block.getId());
			int left = block.getLeft();
			int top = block.getTop();
			int right = block.getRight();
			int bottom = block.getBottom();

			// don't check range to ignore update case,
			// because I still need to sync merge cell data to client side

			result.setData("left", left);
			result.setData("top", top);
			result.setData("right", right);
			result.setData("bottom", bottom);

			HeaderPositionHelper helper = Spreadsheet.this
					.getColumnPositionHelper(sheet);
			final int w = helper.getStartPixel(block.getRight() + 1) - helper.getStartPixel(block.getLeft());
			result.setData("width", w);

			HeaderPositionHelper rhelper = Spreadsheet.this
					.getRowPositionHelper(sheet);
			final int h = rhelper.getStartPixel(block.getBottom() + 1) - rhelper.getStartPixel(block.getTop());
			result.setData("height", h);

			/**
			 * merge_ -> mergeCell
			 */
			response("mergeCell" + Utils.nextUpdateId(), new AuMergeCell(Spreadsheet.this, "", Utils.getSheetUuid(sheet), result.toString()));
		}

		public void addMergeCell(Worksheet sheet, int left, int top, int right,	int bottom) {
			MergeMatrixHelper mmhelper = this.getMergeMatrixHelper(sheet);

			Set toadd = new HashSet();
			Set torem = new HashSet();
			mmhelper.addMergeRange(left, top, right, bottom, toadd, torem);
			for (Iterator iter = torem.iterator(); iter.hasNext();) {
				MergedRect rect = (MergedRect) iter.next();
				log.debug("(A)remove merge:" + rect);
				updateMergeCell0(sheet, rect, "remove");
			}
			for (Iterator iter = toadd.iterator(); iter.hasNext();) {
				MergedRect rect = (MergedRect) iter.next();
				log.debug("add merge:" + rect);
				updateMergeCell0(sheet, rect, "add");
			}
//			updateCell(sheet, left > 0 ? left - 1 : 0, top > 1 ? top - 1 : 0, right + 1, bottom + 1);
			updateCell(sheet, left, top, right, bottom);
		}

		//in pixel
		public void setColumnWidth(Worksheet sheet, int col, int width, int id, boolean hidden) {
			JSONObject result = new JSONObject();
			result.put("type", "column");
			result.put("column", col);
			result.put("width", width);
			result.put("id", id);
			result.put("hidden", hidden);
			smartUpdate("columnSize", (Object) new Object[] { "", Utils.getSheetUuid(sheet), result}, true);
		}

		//in pixels
		public void setRowHeight(Worksheet sheet, int row, int height, int id, boolean hidden) {
			JSONObject result = new JSONObject();
			result.put("type", "row");
			result.put("row", row);
			result.put("height", height);
			result.put("id", id);
			result.put("hidden", hidden);
			smartUpdate("rowSize", (Object) new Object[] { "", Utils.getSheetUuid(sheet), result}, true);
		}

		@Override
		public Boolean getLeftHeaderHiddens(int row) {
			Worksheet sheet = getSelectedSheet();
			HeaderPositionHelper rowHelper = Spreadsheet.this
					.getRowPositionHelper(sheet);
			HeaderPositionInfo info = rowHelper.getInfo(row);
			return info == null ? Boolean.FALSE : Boolean.valueOf(info.hidden);
		}

		@Override
		public Boolean getTopHeaderHiddens(int col) {
			Worksheet sheet = getSelectedSheet();
			HeaderPositionHelper colHelper = Spreadsheet.this
					.getColumnPositionHelper(sheet);
			HeaderPositionInfo info = colHelper.getInfo(col);
			return info == null ? Boolean.FALSE : Boolean.valueOf(info.hidden);
		}
	}

	public void invalidate() {
		super.invalidate();
		doInvalidate();
	}

	/**
	 * Retrieve client side spreadsheet focus.The cell focus and selection will
	 * keep at last status. It is useful if you want get focus back to
	 * spreadsheet after do some outside processing, for example after user
	 * click a outside button or menu item.
	 */
	public void focus() {
		// retrieve focus should work when spreadsheet init or after invalidate.
		// so I use response to implement it.
		JSONObj result = new JSONObj();
		result.setData("type", "retrive");
		
		/**
		 * rename zssfocus -> doRetrieveFocusCmd
		 */
		response("retrieveFocus" + this.getUuid(), new AuRetrieveFocus(this, result.toString()));
	}

	/**
	 * Retrieve client side spreadhsheet focus, move cell focus to position
	 * (row,column) and also scroll the cell to into visible view.
	 * 
	 * @param row row of cell to move
	 * @param column column of cell to move
	 */
	public void focusTo(int row, int column) {
		JSONObj result = new JSONObj();
		result.setData("row", row);
		result.setData("column", column);
		result.setData("type", "moveto");

		/**
		 * rename zssfocusto -> cellFocus
		 */
		response("cellFocusTo" + this.getUuid(), new AuCellFocusTo(this, result.toString()));

		_focusRect.setLeft(column);
		_focusRect.setRight(column);
		_focusRect.setTop(row);
		_focusRect.setBottom(row);
		_selectionRect.setLeft(column);
		_selectionRect.setRight(column);
		_selectionRect.setTop(row);
		_selectionRect.setBottom(row);
	}

	private String getSheetDefaultRules() {

		Worksheet sheet = getSelectedSheet();

		HeaderPositionHelper colHelper = this.getColumnPositionHelper(sheet);
		HeaderPositionHelper rowHelper = this.getRowPositionHelper(sheet);
		MergeMatrixHelper mmhelper = this.getMergeMatrixHelper(sheet);

		boolean hiderow = isHiderowhead();
		boolean hidecol = isHidecolumnhead();
		boolean showgrid = sheet.isDisplayGridlines();

		int th = hidecol ? 1 : this.getTopheadheight();
		int lw = hiderow ? 1 : this.getLeftheadwidth();
		int cp = this._cellpadding;//
		int rh = this.getRowheight();
		int cw = this.getColumnwidth();
		int lh = 20;// default line height;

		if (lh > rh) {
			lh = rh;
		}

		String sheetid = getUuid();
		String name = "#" + sheetid;

		int cellwidth;// default
		int cellheight;// default
		Execution exe = Executions.getCurrent();

		boolean isGecko = exe.isGecko();
		boolean isIE = exe.isExplorer();
		// boolean isIE7 = exe.isExplorer7();

		if (isGecko) {// firefox
			cellwidth = cw;
			cellheight = rh;
		} else {
			cellwidth = cw - 2 * cp - 1;// 1 is border width
			cellheight = rh - 1;// 1 is border width
		}

		int celltextwidth = cw - 2 * cp - 1;// 1 is border width

		StringBuffer sb = new StringBuffer();

		// zcss.setRule(name+" .zsdata",["padding-top","padding-left"],[th+"px",lw+"px"],true,sid);
		sb.append(name).append(" .zsdata{");
		sb.append("padding-top:").append(th).append("px;");
		sb.append("padding-left:").append(lw).append("px;");
		sb.append("}");

		// zcss.setRule(name+" .zsrow","height",rh+"px",true,sid);
		sb.append(name).append(" .zsrow{");
		sb.append("height:").append(rh).append("px;");
		sb.append("}");

		// zcss.setRule(name+" .zscell",["padding","height","width","line-height"],["0px "+cp+"px 0px "+cp+"px",cellheight+"px",cellwidth+"px",lh+"px"],true,sid);
		sb.append(name).append(" .zscell{");
		sb.append("padding:").append("0px " + cp + "px 0px " + cp + "px;");
		sb.append("height:").append(cellheight).append("px;");
		sb.append("width:").append(cellwidth).append("px;");
		if (!showgrid) {
			sb.append("border-bottom:1px solid #FFFFFF;")
			  .append("border-right:1px solid #FFFFFF;");
		}
		// sb.append("line-height:").append(lh).append("px;\n");
		sb.append("}");

		// zcss.setRule(name+" .zscelltxt",["width","height"],[celltextwidth+"px",cellheight+"px"],true,sid);
		sb.append(name).append(" .zscelltxt{");
		sb.append("width:").append(celltextwidth).append("px;");
		sb.append("height:").append(cellheight).append("px;");
		sb.append("}");

		// zcss.setRule(name+" .zstop",["left","height","line-height"],[lw+"px",(th-2)+"px",lh+"px"],true,sid);

		int toph = th - 1;
		int topheadh = (isGecko) ? toph : toph - 1;
		int cornertoph = th - 1;
		// int topblocktop = toph;
		int fzr = getRowfreeze();
		int fzc = getColumnfreeze();

		if (fzr > -1) {
			toph = toph + rowHelper.getStartPixel(fzr + 1);
		}

		sb.append(name).append(" .zstop{");
		sb.append("left:").append(lw).append("px;");
		sb.append("height:").append(fzr > -1 ? toph - 1 : toph).append("px;");
		// sb.append("line-height:").append(toph).append("px;\n");
		sb.append("}");

		sb.append(name).append(" .zstopi{");
		sb.append("height:").append(toph).append("px;");
		sb.append("}");

		sb.append(name).append(" .zstophead{");
		sb.append("height:").append(topheadh).append("px;");
		sb.append("}");

		sb.append(name).append(" .zscornertop{");
		sb.append("left:").append(lw).append("px;");
		sb.append("height:").append(cornertoph).append("px;");
		sb.append("}");

		// relative, so needn't set top position.
		/*
		 * sb.append(name).append(" .zstopblock{\n");
		 * sb.append("top:").append(topblocktop).append("px;\n");
		 * sb.append("}\n");
		 */

		// zcss.setRule(name+" .zstopcell",["padding","height","width","line-height"],["0px "+cp+"px 0px "+cp+"px",th+"px",cellwidth+"px",lh+"px"],true,sid);
		sb.append(name).append(" .zstopcell{");
		sb.append("padding:").append("0px " + cp + "px 0px " + cp + "px;");
		sb.append("height:").append(topheadh).append("px;");
		sb.append("width:").append(cellwidth).append("px;");
		sb.append("line-height:").append(topheadh).append("px;");
		sb.append("}");

		// zcss.setRule(name+" .zstopcelltxt","width", celltextwidth
		// +"px",true,sid);
		sb.append(name).append(" .zstopcelltxt{");
		sb.append("width:").append(celltextwidth).append("px;");
		sb.append("}");

		// zcss.setRule(name+" .zsleft",["top","width"],[th+"px",(lw-2)+"px"],true,sid);

		int leftw = lw - 1;
		int leftheadw = leftw;
		int leftblockleft = leftw;

		if (fzc > -1) {
			leftw = leftw + colHelper.getStartPixel(fzc + 1);
		}

		sb.append(name).append(" .zsleft{");
		sb.append("top:").append(th).append("px;");
		sb.append("width:").append(fzc > -1 ? leftw - 1 : leftw)
				.append("px;");
		sb.append("}");

		sb.append(name).append(" .zslefti{");
		sb.append("width:").append(leftw).append("px;");
		sb.append("}");

		sb.append(name).append(" .zslefthead{");
		sb.append("width:").append(leftheadw).append("px;");
		sb.append("}");

		sb.append(name).append(" .zsleftblock{");
		sb.append("left:").append(leftblockleft).append("px;");
		sb.append("}");

		sb.append(name).append(" .zscornerleft{");
		sb.append("top:").append(th).append("px;");
		sb.append("width:").append(leftheadw).append("px;");
		sb.append("}");

		// zcss.setRule(name+" .zsleftcell",["height","line-height"],[(rh-1)+"px",(rh)+"px"],true,sid);//for
		// middle the text, i use row leight instead of lh
		sb.append(name).append(" .zsleftcell{");
		sb.append("height:").append(rh - 1).append("px;");
		sb.append("line-height:").append(rh - 1).append("px;");
		sb.append("}");

		// zcss.setRule(name+" .zscorner",["width","height"],[(lw-2)+"px",(th-2)+"px"],true,sid);

		sb.append(name).append(" .zscorner{");
		sb.append("width:").append(fzc > -1 ? leftw : leftw + 1)
				.append("px;");
		sb.append("height:").append(fzr > -1 ? toph : toph + 1).append("px;");
		sb.append("}");

		sb.append(name).append(" .zscorneri{");
		sb.append("width:").append(lw - 2).append("px;");
		sb.append("height:").append(th - 2).append("px;");
		sb.append("}");

		sb.append(name).append(" .zscornerblock{");
		sb.append("left:").append(lw).append("px;");
		sb.append("top:").append(th).append("px;");
		sb.append("}");

		// zcss.setRule(name+" .zshboun","height",th+"px",true,sid);
		sb.append(name).append(" .zshboun{");
		sb.append("height:").append(th).append("px;");
		sb.append("}");

		// zcss.setRule(name+" .zshboun","height",th+"px",true,sid);
/*		sb.append(name).append(" .zshboun{\n");
		sb.append("height:").append(th).append("px;\n");
		sb.append("}\n");
*/
		// zcss.setRule(name+" .zshbouni","height",th+"px",true,sid);
		sb.append(name).append(" .zshbouni{");
		sb.append("height:").append(th).append("px;");
		sb.append("}");

		sb.append(name).append(" .zsfztop{");
		sb.append("border-bottom-style:").append(fzr > -1 ? "solid" : "none")
				.append(";");
		sb.append("}");
		sb.append(name).append(" .zsfzcorner{");
		sb.append("border-bottom-style:").append(fzr > -1 ? "solid" : "none")
				.append(";");
		sb.append("}");

		sb.append(name).append(" .zsfzleft{");
		sb.append("border-right-style:").append(fzc > -1 ? "solid" : "none")
				.append(";");
		sb.append("}");
		sb.append(name).append(" .zsfzcorner{");
		sb.append("border-right-style:").append(fzc > -1 ? "solid" : "none")
				.append(";");
		sb.append("}");

		// TODO transparent border mode

		boolean transparentBorder = false;

		if (transparentBorder) {
			sb.append(name).append(" .zscell {");
			sb.append("border-right-color: transparent;");
			sb.append("border-bottom-color: transparent;");
			if (isIE) {
				/** for IE6 **/
				String color_to_transparent = "tomato";
				sb.append("_border-color:" + color_to_transparent + ";");
				sb.append("_filter:chroma(color=" + color_to_transparent + ");");
			}
			sb.append("}");
		}

		List<HeaderPositionInfo> infos = colHelper.getInfos();
		for (HeaderPositionInfo info : infos) {
			boolean hidden = info.hidden;
			int index = info.index;
			int width = hidden ? 0 : info.size;
			int cid = info.id;

			celltextwidth = width - 2 * cp - 1;// 1 is border width

			// bug 1989680
			if (celltextwidth < 0)
				celltextwidth = 0;

			if (!isGecko) {
				cellwidth = celltextwidth;
			} else {
				cellwidth = width;
			}

			if (width <= 0) {
				sb.append(name).append(" .zsw").append(cid).append("{");
				sb.append("display:none;");
				sb.append("}");

			} else {
				sb.append(name).append(" .zsw").append(cid).append("{");
				sb.append("width:").append(cellwidth).append("px;");
				sb.append("}");

				sb.append(name).append(" .zswi").append(cid).append("{");
				sb.append("width:").append(celltextwidth).append("px;");
				sb.append("}");
			}
		}

		infos = rowHelper.getInfos();
		for (HeaderPositionInfo info : infos) {
			boolean hidden = info.hidden;
			int index = info.index;
			int height = hidden ? 0 : info.size;
			int cid = info.id;
			cellheight = height;
			if (!isGecko) {
				cellheight = height - 1;// 1 is border width
			}

			if (height <= 0) {

				sb.append(name).append(" .zsh").append(cid).append("{\n");
				sb.append("display:none;");
				sb.append("}");

				sb.append(name).append(" .zslh").append(cid).append("{");
				sb.append("display:none;");
				sb.append("}");

			} else {
				sb.append(name).append(" .zsh").append(cid).append("{");
				sb.append("height:").append(height).append("px;");
				sb.append("}");

				sb.append(name).append(" .zshi").append(cid).append("{");
				sb.append("height:").append(cellheight).append("px;");
				sb.append("}");

				int h2 = (height < 1) ? 0 : height - 1;

				sb.append(name).append(" .zslh").append(cid).append("{");
				sb.append("height:").append(h2).append("px;");
				sb.append("line-height:").append(h2).append("px;");
				sb.append("}");

			}
		}
		sb.append(".zs_header{}");// for indicating add new rule before this

		// merge size;
		List ranges = mmhelper.getRanges();
		Iterator iter = ranges.iterator();
		final int defaultSize = colHelper.getDefaultSize();
		final int defaultRowSize = rowHelper.getDefaultSize();

		while (iter.hasNext()) {
			MergedRect block = (MergedRect) iter.next();
			int left = block.getLeft();
			int right = block.getRight();
			int width = 0;
			for (int i = left; i <= right; i++) {
				final HeaderPositionInfo info = colHelper.getInfo(i);
				if (info != null) {
					final boolean hidden = info.hidden;
					final int colSize = hidden ? 0 : info.size;
					width += colSize;
				} else {
					width += defaultSize ;
				}
			}
			int top = block.getTop();
			int bottom = block.getBottom();
			int height = 0;
			for (int i = top; i <= bottom; i++) {
				final HeaderPositionInfo info = rowHelper.getInfo(i);
				if (info != null) {
					final boolean hidden = info.hidden;
					final int rowSize = hidden ? 0 : info.size;
					height += rowSize;
				} else {
					height += defaultRowSize ;
				}
			}

			if (width <= 0 || height <= 0) { //total hidden
				sb.append(name).append(" .zsmerge").append(block.getId()).append("{");
				sb.append("display:none;");
				sb.append("}");

				sb.append(name).append(" .zsmerge").append(block.getId());
				sb.append(" .zscelltxt").append("{");
				sb.append("display:none;");
				sb.append("}");
			} else {
				celltextwidth = width - 2 * cp - 1;// 1 is border width
				int celltextheight = height - 1; //1 is border height
	
				if (!isGecko) {
					cellwidth = celltextwidth;
					cellheight = celltextheight;
				} else {
					cellwidth = width;
					cellheight = height;
				}
				
				sb.append(name).append(" .zsmerge").append(block.getId()).append("{");
				sb.append("width:").append(cellwidth).append("px;");
				sb.append("height:").append(cellheight).append("px;");
				sb.append("}");
	
				sb.append(name).append(" .zsmerge").append(block.getId());
				sb.append(" .zscelltxt").append("{");
				sb.append("width:").append(celltextwidth).append("px;");
				sb.append("height:").append(celltextheight).append("px;");
				sb.append("}");
			}
		}

		sb.append(".zs_indicator{}");// for indicating the css is load ready

		return sb.toString();
	}

	/**
	 * Returns the encoded URL for the dynamic generated content, or empty the
	 * component doesn't belong to any desktop.
	 */
	private static String getDynamicMediaURI(AbstractComponent comp, int version, String name, String format) {
		final Desktop desktop = comp.getDesktop();
		if (desktop == null)
			return ""; // no avail at client

		final StringBuffer sb = new StringBuffer(64).append('/');
		Strings.encode(sb, version);
		if (name != null || format != null) {
			sb.append('/');
			boolean bExtRequired = true;
			if (name != null && name.length() != 0) {
				sb.append(name.replace('\\', '/'));
				bExtRequired = name.lastIndexOf('.') < 0;
			} else {
				sb.append(comp.getUuid());
			}
			if (bExtRequired && format != null)
				sb.append('.').append(format);
		}
		return desktop.getDynamicMediaURI(comp, sb.toString()); // already
		// encoded
	}

	private int getDataPanelWidth() {
		int col = getMaxcolumns();

		HeaderPositionHelper colHelper = getColumnPositionHelper(getSelectedSheet());

		return colHelper.getStartPixel(col);
	}

	private int getDataPanelHeight() {
		int row = getMaxrows();

		HeaderPositionHelper rowHelper = getRowPositionHelper(getSelectedSheet());

		return rowHelper.getStartPixel(row);
	}

	private void doSheetClean(Worksheet sheet) {
		if (getBook().getSheetIndex(sheet) != -1)
			deleteFocus();
		List list = loadWidgetLoaders();
		int size = list.size();
		for (int i = 0; i < size; i++) {
			((WidgetLoader) list.get(i)).onSheetClean(sheet);
		}
		removeAttribute(MERGE_MATRIX_KEY);
		_loadedRect.set(-1, -1, -1, -1);
		_selectionRect.set(0, 0, 0, 0);
		_focusRect.set(0, 0, 0, 0);
		
		//issue #99: freeze column and freeze row is not reset. 
		_colFreeze = -1;
		_colFreezeset = false;
		_rowFreeze = -1;
		_rowFreezeset = false;
	}

	private void doSheetSelected(Worksheet sheet) {
		//load widgets
		List list = loadWidgetLoaders();
		int size = list.size();
		for (int i = 0; i < size; i++) {
			((WidgetLoader) list.get(i)).onSheetSelected(sheet);
		}
		//setup gridline
		setDisplayGridlines(_selectedSheet.isDisplayGridlines());
		setProtectSheet(_selectedSheet.getProtect());
		//register collaborated focus
		moveFocus();
		_selectedSheetName = _selectedSheet.getSheetName();
	}
	
	public String getSelectedSheetName() {
		return _selectedSheet != null ? _selectedSheetName : null;
	}
	
	private void addChartWidget(Worksheet sheet, ZssChartX chart) {
		//load widgets
		List list = loadWidgetLoaders();
		int size = list.size();
		for (int i = 0; i < size; i++) {
			((WidgetLoader) list.get(i)).addChartWidget(sheet, chart);
		}
	}

	private void addPictureWidget(Worksheet sheet, Picture picture) {
		//load widgets
		List list = loadWidgetLoaders();
		int size = list.size();
		for (int i = 0; i < size; i++) {
			((WidgetLoader) list.get(i)).addPictureWidget(sheet, picture);
		}
	}

	private void deletePictureWidget(Worksheet sheet, Picture picture) {
		//load widgets
		List list = loadWidgetLoaders();
		int size = list.size();
		for (int i = 0; i < size; i++) {
			((WidgetLoader) list.get(i)).deletePictureWidget(sheet, picture);
		}
	}
	
	private void updatePictureWidget(Worksheet sheet, Picture picture) {
		//load widgets
		List list = loadWidgetLoaders();
		int size = list.size();
		for (int i = 0; i < size; i++) {
			((WidgetLoader) list.get(i)).updatePictureWidget(sheet, picture);
		}
	}

	private void deleteChartWidget(Worksheet sheet, Chart chart) {
		//load widgets
		List list = loadWidgetLoaders();
		int size = list.size();
		for (int i = 0; i < size; i++) {
			((WidgetLoader) list.get(i)).deleteChartWidget(sheet, chart);
		}
	}

	private void updateChartWidget(Worksheet sheet, Chart chart) {
		//load widgets
		List list = loadWidgetLoaders();
		int size = list.size();
		for (int i = 0; i < size; i++) {
			((WidgetLoader) list.get(i)).updateChartWidget(sheet, chart);
		}
	}

	private void clearHeaderSizeHelper(boolean row, boolean col) {
		if (row)
			removeAttribute(ROW_SIZE_HELPER_KEY);
		if (col)
			removeAttribute(COLUMN_SIZE_HELPER_KEY);
	}

	private static String getSizeHelperStr(HeaderPositionHelper helper) {
		List<HeaderPositionInfo> infos = helper.getInfos();
		StringBuffer csc = new StringBuffer();
		for(HeaderPositionInfo info : infos) {
			if (csc.length() > 0)
				csc.append(",");
			csc.append(info.index).append(",")
				.append(info.size).append(",")
				.append(info.id).append(",")
				.append(info.hidden);
		}
		return csc.toString();
	}

	static private String getRectStr(Rect rect) {
		StringBuffer sb = new StringBuffer();
		sb.append(rect.getLeft()).append(",").append(rect.getTop()).append(",")
				.append(rect.getRight()).append(",").append(rect.getBottom());
		return sb.toString();
	}

	private void doInvalidate() {
		// reset
		_loadedRect.set(-1, -1, -1, -1);
		Worksheet sheet = getSelectedSheet();

		clearHeaderSizeHelper(true, true);
		// remove this, beacuse invalidate will cause client side rebuild,
		// i must reinitial size helper since there are maybe some customized is
		// from client.
		// System.out.println(">>>>>>>>>>>remove this");
		// removeAttribute(MERGE_MATRIX_KEY);//TODO remove this, for insert
		// column test only

		_custColId = new SequenceId(-1, 2);
		_custRowId = new SequenceId(-1, 2);

		this.getWidgetHandler().invaliate();

		List list = loadWidgetLoaders();
		int size = list.size();
		for (int i = 0; i < size; i++) {
			((WidgetLoader) list.get(i)).invalidate();
		}
	}

	/**
	 * load widget loader
	 */
	/* package */List<WidgetLoader> loadWidgetLoaders() {
		if (_widgetLoaders != null)
			return _widgetLoaders;
		_widgetLoaders = new ArrayList<WidgetLoader>();
		final String loaderclzs = (String) Library.getProperty(WIDGET_LOADERS);
		if (loaderclzs != null) {
			try {
				String[] clzs = loaderclzs.split(",");
				WidgetLoader wl;
				for (int i = 0; i < clzs.length; i++) {
					clzs[i] = clzs[i].trim();
					if ("".equals(clzs[i]))
						continue;
					wl = (WidgetLoader) Classes.newInstance(clzs[i], null, null);
					wl.init(this);
					_widgetLoaders.add(wl);
				}
			} catch (Exception x) {
				throw new UiException(x);
			}
		}
		return _widgetLoaders;
	}

	/**
	 * get widget handler
	 */
	private WidgetHandler getWidgetHandler() {
		if (_widgetHandler == null) {
			_widgetHandler = newWidgetHandler();
		}
		return _widgetHandler;
	}

	/**
	 * new widget handler
	 * 
	 * @return
	 */
	private WidgetHandler newWidgetHandler() {
		final String handlerclz = (String) Library.getProperty(WIDGET_HANDLER);
		if (handlerclz != null) {
			try {
				_widgetHandler = (WidgetHandler) Classes.newInstance(handlerclz, null, null);
				_widgetHandler.init(this);
			} catch (Exception x) {
				throw new UiException(x);
			}
		} else {
			_widgetHandler = new VoidWidgetHandler();
			_widgetHandler.init(this);
		}
		return _widgetHandler;
	}

	/**
	 * Add widget to the {@link WidgetHandler} of this spreadsheet,
	 * 
	 * @param widget a widget
	 * @return true if success to add a widget
	 */
	private boolean addWidget(Widget widget) {
		return getWidgetHandler().addWidget(widget);
	}

	/**
	 * Remove widget from the {@link WidgetHandler} of this spreadsheet,
	 * 
	 * @param widget
	 * @return true if success to remove a widget
	 */
	private boolean removeWidget(Widget widget) {
		return getWidgetHandler().removeWidget(widget);
	}
	
	/**
	 * Default widget implementation, don't provide any function.
	 */
	private class VoidWidgetHandler implements WidgetHandler, Serializable {

		public boolean addWidget(Widget widget) {
			return false;
		}

		public Spreadsheet getSpreadsheet() {
			return Spreadsheet.this;
		};

		public void invaliate() {
		}

		public void onLoadOnDemand(Worksheet sheet, int left, int top, int right, int bottom) {
		}

		public boolean removeWidget(Widget widget) {
			return false;
		}

		public void init(Spreadsheet spreadsheet) {
		}
		
		public void updateWidgets(Worksheet sheet, int left, int top, int right, int bottom) {
		}
	}

	private void processStartEditing(String token, StartEditingEvent event, String editingType) {
		if (!event.isCancel()) {
			Object val;
			final boolean useEditValue = event.isEditingSet() || event.getClientValue() == null; 
			if (useEditValue) {
				val = event.getEditingValue();
			} else {
				val = event.getClientValue();
			}

			processStartEditing0(token, event.getSheet(), event.getRow(), event
					.getColumn(), val, useEditValue, editingType);
		} else {
			processCancelEditing0(token, event.getSheet(), event.getRow(),
					event.getColumn(), false, editingType);
		}
	}

	private void processStopEditing(String token, StopEditingEvent event, String editingType) {
		if (!event.isCancel()) {
			processStopEditing0(token, event.getSheet(), event.getRow(), event.getColumn(), event.getEditingValue(), editingType);
		} else
			processCancelEditing0(token, event.getSheet(), event.getRow(), event.getColumn(), false, editingType);
	}
	
	private void showFormulaError(FormulaParseException ex) {
		try {
			Messagebox.show(ex.getMessage(), "ZK Spreadsheet", Messagebox.OK, Messagebox.EXCLAMATION, new EventListener() {
				public void onEvent(Event evt) {
					Spreadsheet.this.focus();
				}
			});
		} catch (InterruptedException e) {
			// ignore
		}
	}

	private void processStopEditing0(final String token, final Worksheet sheet, final int rowIdx, final int colIdx, final Object value, final String editingType) {
		try {
			if(!Utils.setEditTextWithValidation(this, sheet, rowIdx, colIdx, value == null ? "" : value.toString(),
				//callback
				new EventListener() {
					@Override
					public void onEvent(Event event) throws Exception {
						final String eventname = event.getName();
						if (Messagebox.ON_CANCEL.equals(eventname)) { //cancel
							Spreadsheet.this.processCancelEditing0(token, sheet, rowIdx, colIdx, true, editingType); //skipMove
						} else if (Messagebox.ON_OK.equals(eventname)) { //ok
							Spreadsheet.this.processStopEditing0(token, sheet, rowIdx, colIdx, value, editingType);
						} else { //retry
							Spreadsheet.this.processRetryEditing0(token, sheet, rowIdx, colIdx, value, editingType);
						}
					}
				}
			)) {
				return;
			}

			//JSONObj result = new JSONObj();
			JSONObject result = new JSONObject();
			result.put("r", rowIdx);
			result.put("c", colIdx);
			result.put("type", "stopedit");
			result.put("val", "");
			result.put("et", editingType);

			smartUpdate("dataUpdateStop", new Object[] { token,	Utils.getSheetUuid(sheet), result});
		} catch (RuntimeException x) {
			processCancelEditing0(token, sheet, rowIdx, colIdx, false, editingType);
			if (x instanceof FormulaParseException) {
				showFormulaError((FormulaParseException)x);
			} else {
				throw x;
			}
		}
	}

	private void processStartEditing0(String token, Worksheet sheet, int row, int col, Object value, boolean useEditValue, String editingType) {
		try {
			JSONObject result = new JSONObject();
			result.put("r", row);
			result.put("c", col);
			result.put("type", "startedit");
			result.put("val", value == null ? "" : value.toString());
			result.put("et", editingType);
			if (useEditValue) { //shall use edit value from server
				result.put("server", true); 
			}
			smartUpdate("dataUpdateStart", new Object[] { token, Utils.getSheetUuid(sheet), result});
		} catch (RuntimeException x) {
			processCancelEditing0(token, sheet, row, col, false, editingType);
			throw x;
		}
	}

	private void processCancelEditing0(String token, Worksheet sheet, int row, int col, boolean skipMove, String editingType) {
		JSONObject result = new JSONObject();
		result.put("r", row);
		result.put("c", col);
		result.put("type", "canceledit");
		result.put("val", "");
		result.put("sk", skipMove);
		result.put("et", editingType);
		smartUpdate("dataUpdateCancel", new Object[] { token, Utils.getSheetUuid(sheet), result});
	}

	private void processRetryEditing0(String token, Worksheet sheet, int row, int col, Object value, String editingType) {
		try {
			processCancelEditing0(token, sheet, row, col, true, editingType);
			JSONObject result = new JSONObject();
			result.put("r", row);
			result.put("c", col);
			result.put("type", "retryedit");
			result.put("val", value);
			result.put("et", editingType);
			smartUpdate("dataUpdateRetry", new Object[] { "", Utils.getSheetUuid(sheet), result});
		} catch (RuntimeException x) {
			processCancelEditing0(token, sheet, row, col, false, editingType);
			throw x;
		}
	}

	public boolean insertBefore(Component newChild, Component refChild) {
		// not all child can insert into spreadsheet a child want to insert to spreadsheet must provide some speaciall
		// attribute
		if (newChild.getAttribute(SpreadsheetCtrl.CHILD_PASSING_KEY) != null) {
			return super.insertBefore(newChild, refChild);
		} else {
			throw new UiException("Unsupported child for Spreadsheet: " + newChild);
		}
	}
	/**
	 * push the current cell state of selection region and selected sheet
	 */
/*	public void pushCellState(){
		stateManager.pushCellState();
	}
*/	
	/**
	 * push current cell state of selected sheet and specified region rect
	 * @param rect
	 */
/*	public void pushCellState(Rect rect){
		stateManager.pushCellState(rect);
	}
*/	
	/**
	 * push current cell state in specified sheets and rects
	 * @param sheets
	 * @param rectArray
	 */
/*	public void pushCellState(Sheet[] sheets, Rect[] rectArray){
		stateManager.pushCellState(sheets, rectArray);
	}
*/	
	/**
	 * push certain state to redostack
	 * @param iState
	 */
/*	public void pushRedoState(IState iState){
		stateManager.pushRedoState(iState);
	}
*/	
	/**
	 * 
	 * @return the top IState in the undostack
	 */
/*	public IState peekUndoState(){
		return stateManager.peekUndoStack();
	}
*/	
	/**
	 * push the current state into undostack before change the col/row header size
	 * @param left
	 * @param top
	 * @param right
	 * @param bottom
	 */
/*	public void pushRowColSizeState(int left, int top, int right, int bottom){	
		stateManager.pushRowColSizeState(left, top, right, bottom);
	}
*/	
	/**
	 * push the current state into undostack before insert the row/col operation
	 * @param left
	 * @param top
	 * @param right
	 * @param bottom
	 */
	
/*	public void pushInsertRowColState(int left, int top, int right, int bottom){
		stateManager.pushInsertRowColState(left, top, right, bottom);
	}
*/	
	/**
	 * push the current state into undostack before delete row/col operation
	 * @param left
	 * @param top
	 * @param right
	 * @param bottom
	 */
/*	public void pushDeleteRowColState(int left, int top, int right, int bottom){
		stateManager.pushDeleteRowColState(left, top, right, bottom);
	}
*/	
	/**
	 * undo operation and notify other editors
	 */
/*	public void undo(){
		stateManager.undo();
	}
*/	
	/**
	 * redo, and notify other editors 
	 */
/*	public void redo(){
		stateManager.redo();
	}
*/	
	/**
	 * return the undo/redo state manager
	 */
/*	public StateManager getStateManager(){
		return this.stateManager;
	}
*/	
	/**
	 * Remove editor's focus on specified name
	 */
	public void removeEditorFocus(String id){
		response("removeEditorFocus" + _focusId.next(), new AuInvoke((Component)this, "removeEditorFocus", id));
		_focuses.remove(id);
	}
	
	/**
	 *  Add and move other editor's focus
	 */
	public void moveEditorFocus(String id, String name, String color, int row ,int col){
		if (_focus != null && !_focus.id.equals(id)) {
			response("moveEditorFocus" + _focusId.next(), new AuInvoke((Component)this, "moveEditorFocus", new String[]{id, name, color,""+row,""+col}));
			_focuses.put(id, new Focus(id, name, color, row, col, null));
		}
	}
	
	private void syncEditorFocus() {
		if (_book != null) {
			synchronized(_book) {
				for(final Iterator<Focus> it = _focuses.values().iterator(); it.hasNext();) {
					final Focus focus = it.next();
					if (!((BookCtrl)_book).containsFocus(focus)) { //
						it.remove();
						removeEditorFocus(focus.id); //remove from the client
					}
				}
			}
		}
	}
	
	/**
	 * Set focus name of this spreadsheet.
	 * @param name focus name that show on other Spreadsheet
	 */
	public void setUserName(String name) {
		if (_focus != null) {
			_focus.name = name;
		}
		_userName = name;
	}

	//sync friend focus position
	private void syncFriendFocusesPosition(int left, int top, int right, int bottom) {
		int row = -1, col = -1;
		for(Focus focus : _focuses.values()) {
			row=focus.row;
			col=focus.col;
			if(col>=left && col<=right && row>=top  && row<=bottom) {
				this.moveEditorFocus(focus.id, focus.name, focus.color, row, col);
			}
		}
	}
	
	/**
	 * update/invalidate all focus/selection/hightlight to align with cell border
	 */
	public void updateFocus(int left, int top, int right, int bottom){
		int row,col,sL,sT,sR,sB,hL,hT,hR,hB;
		row=col=sL=sT=sR=sB=hL=hT=hR=hB=-1;
		Position pos = this.getCellFocus();
		if(pos!=null){
			row=pos.getRow();
			col=pos.getColumn();
			response("updateSelfFocus", new AuInvoke((Component)this,"updateSelfFocus", new String[]{""+row,""+col}));
		}
		Rect rect = this.getSelection();
		if(rect!=null){
			sL=rect.getLeft();
			sT=rect.getTop();
			sR=rect.getRight();
			sB=rect.getBottom();	
			response("updateSelfSelection", new AuInvoke((Component)this,"updateSelfSelection", new String[]{""+sL,""+sT,""+sR,""+sB}));
		}
		
		rect=this.getHighlight();
		if(rect!=null){
			hL=rect.getLeft();
			hT=rect.getTop();
			hR=rect.getRight();
			hB=rect.getBottom();
			
			response("updateSelfHightlight", new AuInvoke((Component)this,"updateSelfHighlight", new String[]{""+hL,""+hT,""+hR,""+hB}));
			
		}
	}
	
	/**
	 * @param sheet
	 */
	//it will be call when delete sheet 
/*	public void cleanRelatedState(Sheet sheet){
		stateManager.cleanRelatedState(sheet);
	}
*/

	static {
		addClientEvent(Spreadsheet.class, Events.ON_CELL_FOUCSED, CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, Events.ON_CELL_SELECTION,	CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, Events.ON_SELECTION_CHANGE, CE_IMPORTANT | CE_DUPLICATE_IGNORE | CE_NON_DEFERRABLE);
		addClientEvent(Spreadsheet.class, Events.ON_CELL_CLICK, CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, Events.ON_CELL_RIGHT_CLICK, CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, Events.ON_CELL_DOUBLE_CLICK, CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, Events.ON_HEADER_CLICK, CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, Events.ON_HEADER_RIGHT_CLICK,	CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, Events.ON_HEADER_DOUBLE_CLICK, CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, Events.ON_HYPERLINK, 0);
		addClientEvent(Spreadsheet.class, Events.ON_INSERT_FORMULA, 0);
		addClientEvent(Spreadsheet.class, Events.ON_FILTER, CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, Events.ON_VALIDATE_DROP, CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, org.zkoss.zk.ui.event.Events.ON_CTRL_KEY, CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, org.zkoss.zk.ui.event.Events.ON_BLUR,	CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		
		addClientEvent(Spreadsheet.class, Events.ON_START_EDITING, 0);
		addClientEvent(Spreadsheet.class, Events.ON_EDITBOX_EDITING, 0);
		addClientEvent(Spreadsheet.class, Events.ON_STOP_EDITING, 0);

		addClientEvent(Spreadsheet.class, "onZSSCellFocused", CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, "onZSSCellFetch", CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, "onZSSSyncBlock", CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, "onZSSHeaderModif", CE_IMPORTANT);
		addClientEvent(Spreadsheet.class, "onZSSCellMouse", CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, "onZSSHeaderMouse", CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, "onZSSMoveWidget", CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, "onZSSWidgetCtrlKey", CE_IMPORTANT | CE_DUPLICATE_IGNORE);
	}

	final Command[] Commands = { new BlockSyncCommand(),
			new CellFetchCommand(), new CellSelectionCommand(),
			new CellFocusedCommand(), new CellMouseCommand(),
			new SelectionChangeCommand(), new HeaderMouseCommand(),
			new HeaderCommand(), new StartEditingCommand(),
			new StopEditingCommand(), new EditboxEditingCommand(), 
			new MoveWidgetCommand(), new WidgetCtrlKeyCommand(), new FilterCommand()};

	// super//
	/**
	 * Processes an AU request. It is invoked internally.
	 * 
	 * <p>
	 * Default: in addition to what are handled by {@link XulElement#service},
	 * it also handles onChange.
	 */
	public void service(AuRequest request, boolean everError) {
		final String cmd = request.getCommand();

		if (Events.ON_HYPERLINK.equals(cmd)) {
			final HyperlinkEvent evt = HyperlinkEvent.getHyperlinkEvent(request);
			if (evt != null) {
				org.zkoss.zk.ui.event.Events.postEvent(evt);
			}
			return;
		}
		for (Command comObj : Commands) {
			if (comObj.getCommand().equals(cmd)) {
				comObj.process(request);
				return;
			}
		}
		
		super.service(request, everError);
	}
	
	public void smartUpdate(String attr, Object value) {
		super.smartUpdate(attr, value);
	}
	
	public void smartUpdate(String attr, Object value, boolean append) {
		super.smartUpdate(attr, value, append);
	}
	
	public void response(String key, AuResponse response) {
		super.response(key, response);
	}
	
	private final static Map<Integer, String> iconMap = new HashMap<Integer, String>(3);
	static {
		iconMap.put(ErrorStyle.INFO, Messagebox.INFORMATION);
		iconMap.put(ErrorStyle.STOP, Messagebox.ERROR);
		iconMap.put(ErrorStyle.WARNING, Messagebox.EXCLAMATION);
	}
	//return true if a valid input; false otherwise and show Error Alert if required 
	public boolean validate(Worksheet sheet, final int row, final int col, final String txt, 
		final EventListener callback) {
		final Worksheet ssheet = this.getSelectedSheet();
		if (ssheet == null || !ssheet.equals(sheet)) { //skip no sheet case
			return true;
		}
		if (_inCallback) { //skip validation check
			return true;
		}
		final Range rng = Ranges.range(sheet, row, col);
		final DataValidation dv = rng.validate(txt);
		if (dv != null) {
			if (dv.getShowErrorBox()) {
				String errTitle = dv.getErrorBoxTitle();
				String errText = dv.getErrorBoxText();
				if (errTitle == null) {
					errTitle = "ZK Spreadsheet";
				}
				if (errText == null) {
					errText = "The value you entered is not valid.\n\nA user has restricted values that can be entered into this cell.";
				}
				final int errStyle = dv.getErrorStyle();
				try {
					switch(errStyle) {
						case ErrorStyle.STOP:
						{
							final int btn = Messagebox.show(
								errText, errTitle, Messagebox.RETRY|Messagebox.CANCEL, 
								Messagebox.ERROR, Messagebox.RETRY, new EventListener() {
								@Override
								public void onEvent(Event event) throws Exception {
									final String evtname = event.getName();
									if (Messagebox.ON_RETRY.equals(evtname)) {
										retry(callback);
									} else if (Messagebox.ON_CANCEL.equals(evtname)) {
										cancel(callback);
									}
								}
							});
						}
						break;
						case ErrorStyle.WARNING:
						{
							errText += "\n\nContinue?";
							final int btn = Messagebox.show(
								errText, errTitle, Messagebox.YES|Messagebox.NO|Messagebox.CANCEL, 
								Messagebox.EXCLAMATION, Messagebox.NO, new EventListener() {
								@Override
								public void onEvent(Event event) throws Exception {
									final String evtname = event.getName();
									if (Messagebox.ON_NO.equals(evtname)) {
										retry(callback);
									} else if (Messagebox.ON_CANCEL.equals(evtname)) {
										cancel(callback);
									} else if (Messagebox.ON_YES.equals(evtname)) {
										ok(callback);
									}
								}
							});
							if (getDesktop().getWebApp().getConfiguration().isEventThreadEnabled() && btn == Messagebox.YES) {
								return true;
							}
						}
						break;
						case ErrorStyle.INFO:
						{
							final int btn = Messagebox.show(
								errText, errTitle, Messagebox.OK|Messagebox.CANCEL, 
								Messagebox.INFORMATION, Messagebox.OK, new EventListener() {
								@Override
								public void onEvent(Event event) throws Exception {
									final String evtname = event.getName();
									if (Messagebox.ON_CANCEL.equals(evtname)) {
										cancel(callback);
									} else if (Messagebox.ON_OK.equals(evtname)) {
										ok(callback);
									}
								}
							});
							if (getDesktop().getWebApp().getConfiguration().isEventThreadEnabled() && btn == Messagebox.OK) {
								return true;
							}
						}
						break;
					}
				} catch (InterruptedException e) {
					//ignore
				}
			}
			return false;
		}
		return true;
	}
	
	private boolean _inCallback = false;
	private void errorBoxCallback(EventListener callback, String eventname) {
		if (!getDesktop().getWebApp().getConfiguration().isEventThreadEnabled() && callback != null) {
			try {
				_inCallback = true;
				callback.onEvent(new Event(eventname, this));
			} catch (Exception e) {
				throw UiException.Aide.wrap(e);
			} finally {
				_inCallback = false;
			}
		}
	}
	//when user press OK/YES button of the validation ErrorBox, have to call back to resend the setEditText() operation 
	private void ok(EventListener callback) {
		errorBoxCallback(callback, Messagebox.ON_OK);
	}
	//when user press RETRY/NO button of the validation ErrorBox, have to call back to handle UI operation 
	private void retry(EventListener callback) {
		//TODO: shall set focus back to cell at (row, col), select the text, enter edit mode
		errorBoxCallback(callback, Messagebox.ON_RETRY);
	}
	//when user press CANCEL button of the validation ErrorBox, have to call back to handle UI operation 
	private void cancel(EventListener callback) {
		//TODO: shall set focus back to cell at (row, col) and restore cell value
		errorBoxCallback(callback, Messagebox.ON_CANCEL);
	}
}