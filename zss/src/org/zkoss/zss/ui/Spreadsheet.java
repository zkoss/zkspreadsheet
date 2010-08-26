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

import static org.zkoss.zss.ui.fn.UtilFns.getCellInnerAttrs;
import static org.zkoss.zss.ui.fn.UtilFns.getCellOuterAttrs;
import static org.zkoss.zss.ui.fn.UtilFns.getCelltext;
import static org.zkoss.zss.ui.fn.UtilFns.getColBegin;
import static org.zkoss.zss.ui.fn.UtilFns.getColEnd;
import static org.zkoss.zss.ui.fn.UtilFns.getRowBegin;
import static org.zkoss.zss.ui.fn.UtilFns.getRowEnd;
import static org.zkoss.zss.ui.fn.UtilFns.getRowOuterAttrs;
import static org.zkoss.zss.ui.fn.UtilFns.getTopHeaderInnerAttrs;
import static org.zkoss.zss.ui.fn.UtilFns.getTopHeaderOuterAttrs;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.Serializable;
import java.net.URL;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.zkoss.lang.Classes;
import org.zkoss.lang.Objects;
import org.zkoss.lang.Strings;
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
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.ext.render.DynamicMedia;
import org.zkoss.zk.ui.sys.ContentRenderer;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.engine.event.EventDispatchListener;
import org.zkoss.zss.engine.event.SSDataEvent;
import org.zkoss.zss.model.Book;
//import org.zkoss.zss.model.Cell;
//import org.zkoss.zss.model.Format;
import org.zkoss.zss.model.FormatText;
import org.zkoss.zss.model.Importer;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Size;
//import org.zkoss.zss.model.Sheet;
//import org.zkoss.zss.model.Size;
//import org.zkoss.zss.model.TextHAlign;
//import org.zkoss.zss.model.event.SSDataEvent;
//import org.zkoss.zss.model.event.SSDataListener;
//import org.zkoss.zss.model.impl.BookImpl;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.model.impl.ExcelImporter;
//import org.zkoss.zss.model.state.IState;
//import org.zkoss.zss.model.state.StateManager;
import org.zkoss.zss.ui.au.in.BlockSyncCommand;
import org.zkoss.zss.ui.au.in.CellFetchCommand;
import org.zkoss.zss.ui.au.in.CellFocusedCommand;
import org.zkoss.zss.ui.au.in.CellMouseCommand;
import org.zkoss.zss.ui.au.in.CellSelectionCommand;
import org.zkoss.zss.ui.au.in.Command;
import org.zkoss.zss.ui.au.in.EditboxEditingCommand;
import org.zkoss.zss.ui.au.in.HeaderCommand;
import org.zkoss.zss.ui.au.in.HeaderMouseCommand;
import org.zkoss.zss.ui.au.in.SelectionChangeCommand;
import org.zkoss.zss.ui.au.in.StartEditingCommand;
import org.zkoss.zss.ui.au.in.StopEditingCommand;
import org.zkoss.zss.ui.au.out.AuCellFocus;
import org.zkoss.zss.ui.au.out.AuCellFocusTo;
import org.zkoss.zss.ui.au.out.AuDataUpdate;
import org.zkoss.zss.ui.au.out.AuHighlight;
import org.zkoss.zss.ui.au.out.AuInsertRowColumn;
import org.zkoss.zss.ui.au.out.AuMergeCell;
import org.zkoss.zss.ui.au.out.AuRemoveRowColumn;
import org.zkoss.zss.ui.au.out.AuRetrieveFocus;
import org.zkoss.zss.ui.au.out.AuSelection;
import org.zkoss.zss.ui.event.CellMouseEvent;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.event.HyperlinkEvent;
import org.zkoss.zss.ui.event.StartEditingEvent;
import org.zkoss.zss.ui.event.StopEditingEvent;
import org.zkoss.zss.ui.fn.UtilFns;
import org.zkoss.zss.ui.impl.CellFormatHelper;
import org.zkoss.zss.ui.impl.HeaderPositionHelper;
import org.zkoss.zss.ui.impl.JSONObj;
import org.zkoss.zss.ui.impl.MergeMatrixHelper;
import org.zkoss.zss.ui.impl.MergedRect;
import org.zkoss.zss.ui.impl.SequenceId;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zss.ui.impl.HeaderPositionHelper.HeaderPositionInfo;
import org.zkoss.zss.ui.sys.SpreadsheetCtrl;
import org.zkoss.zss.ui.sys.SpreadsheetInCtrl;
import org.zkoss.zss.ui.sys.SpreadsheetOutCtrl;
import org.zkoss.zss.ui.sys.WidgetHandler;
import org.zkoss.zss.ui.sys.WidgetLoader;
import org.zkoss.zul.impl.XulElement;

import org.apache.poi.hssf.util.PaneInformation;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

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

public class Spreadsheet extends XulElement {
	private static final Log log = Log.lookup(Spreadsheet.class);

	private static final long serialVersionUID = 1L;
	private static final String ROW_SIZE_HELPER_KEY = "_rowCellSize";
	private static final String COLUMN_SIZE_HELPER_KEY = "_colCellSize";
	private static final String MERGE_MATRIX_KEY = "_mergeRange";
	private static final String WIDGET_HANDLER_KEY = "org.zkoss.zss.ui.sys.WidgetHandler";
	private static final String WIDGET_LOADERS_KEY = "org.zkoss.zss.ui.sys.WidgetLoader";

	transient private Book _book; // the spreadsheet book

	private String _src; // the src to create an internal book
	transient private Importer _importer; // the spreadsheet importer
	private int _maxRows = 20; // how many row of this spreadsheet
	private int _maxColumns = 10; // how many column of this spreadsheet

	private String _selectedSheetId;
	transient private Sheet _selectedSheet;

	private int _rowFreeze = -1; // how many fixed rows
	private boolean _rowFreezeset = false;
	private int _colFreeze = -1; // how many fixed columns
	private boolean _colFreezeset = false;
	private boolean _hideRowhead; // hide row head
	private boolean _hideColhead; // hide column head*/
	/**
	 * Sam add for zss app
	 */
//	StateManager stateManager = new StateManager(this);
	private Collection _focuses = new LinkedList();

	private Rect _focusRect = new Rect(0, 0, 0, 0);
	private Rect _selectionRect = new Rect(0, 0, 0, 0);
	private Rect _loadedRect = new Rect();
	private Rect _highlightRect = null;

	private WidgetHandler _widgetHandler;

	private List _widgetLoaders;


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
	
	public Spreadsheet() {
		this.addEventListener("onStartEditingImpl", new EventListener() {
			public void onEvent(Event event) throws Exception {
				Object[] data = (Object[]) event.getData();
				processStartEditing((String) data[0],
						(StartEditingEvent) data[1]);
			}
		});
		this.addEventListener("onStopEditingImpl", new EventListener() {
			public void onEvent(Event event) throws Exception {

				Object[] data = (Object[]) event.getData();
				processStopEditing((String) data[0], (StopEditingEvent) data[1]);
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
	 * will create a new model depends on url;
	 * 
	 * @return the book model of this spread sheet.
	 */
	public Book getBook() {
		if (_book == null) {
			if (_src == null) {
				throw new UiException("Must specify the src of the spreadsheet to create a new book");
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
						File file = new File(Executions.getCurrent()
								.getDesktop().getWebApp().getRealPath(_src));
						if (file.exists()) {
							url = file.toURI().toURL();
						}
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
	
	private void initBook0(Book book) {
		if (_book != null) {
			_book.unsubscribe(_dataListener);
			_book.removeVariableResolver(_variableResolver);
			_book.removeFunctionMapper(_functionMapper);
		}
		_book = book;
		if (_book != null) {
			_book.subscribe(_dataListener);
			_book.addVariableResolver(_variableResolver);
			_book.addFunctionMapper(_functionMapper);
		}
		if (_selectedSheet != null) {
			doSheetClean(_selectedSheet);
		}
		_selectedSheet = null;
		_selectedSheetId = null;
	}
	
	/**
	 * Sam added for zss app
	 * @param is
	 * @param src
	 */
	public void setBookFromStream(InputStream is, String src){
		if(_selectedSheet!=null){
			doSheetClean(_selectedSheet);
		}
		_selectedSheet=null;
		_selectedSheetId=null;
		setRowfreeze(-1);
		setColumnfreeze(-1);
		//setBook(null);
		_importer = new ExcelImporter();
		_src=src;
		final Book book = ((ExcelImporter)_importer).imports(is, src);
		initBook(book);
	}

	/**
	 * Gets the selected sheet, the default selected sheet is first sheet.
	 * @return #{@link Sheet}
	 */
	public Sheet getSelectedSheet() {
		if (_selectedSheet == null) {
			if (_selectedSheetId == null) {
				if (getBook().getNumberOfSheets() == 0)
					throw new UiException("sheet size of given book is zero");
				_selectedSheet = (Sheet) getBook().getSheetAt(0);
				_selectedSheetId = Utils.getSheetId(_selectedSheet);
			} else {
				_selectedSheet = Utils.getSheetById(_book, _selectedSheetId);
				if (_selectedSheet == null) {
					throw new UiException("can not find sheet by id : "	+ _selectedSheetId);
				}
			}
			doSheetSelected(_selectedSheet);
		}
		return _selectedSheet;
	}

	public Sheet getSheet(int index){
		return (Sheet)getBook().getSheetAt(index);
	}
	
	public int indexOfSheet(Sheet sheet){
		return getBook().getSheetIndex(sheet);
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
		}
	}

	/**
	 * it will change the spreadsheet src name as src
	 * and book name as src => will not reload the book
	 * @param src
	 */
	public void setSrcName(String src) {
		/**
		 * Sam. wait for model integration
		 */
		/*
		BookImpl book = (BookImpl) this.getBook();
		if (book != null)
			book.setName(src);
		_src = src;
		*/
		Book book = (Book) this.getBook();
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
		if (_selectedSheet == null || !_selectedSheet.getSheetName().equals(name)) {
			Sheet sheet = getBook().getSheet(name);
			if (sheet == null) {
				throw new UiException("No such sheet : " + name);
			}

			if (_selectedSheet != null)
				doSheetClean(_selectedSheet);
			_selectedSheet = sheet;
			_selectedSheetId = Utils.getSheetId(_selectedSheet);
			invalidate();
			// the call of onSheetSelected must after invalidate,
			// because i must let invalidate clean lastcellblock first
			doSheetSelected(_selectedSheet);
		}
	}

	/**
	 * Returns the maximum number of rows of this spread sheet. you can assign
	 * new number by calling {@link #setMaxrows(int)}.<br/>
	 * If you do some operation on model, it might increase/decrease the maximum number
	 * of rows automatically such as insert/delete rows into/from the sheet.
	 * 
	 * @return the maximum number of rows.
	 */
	public int getMaxrows() {
		return _maxRows;
	}

	/**
	 * Sets the maximum number of rows of this spread sheet. For example, if you set
	 * to 40, which means it allow row 0 to 39. the minimal value of max number of rows
	 * must large than 0; <br/>
	 * Default : 40.
	 * 
	 * @param maxrows  the maximum number of rows
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

	private String getMaxrowInJSON() {
		JSONObj result = new JSONObj();
		result.setData("maxrow", _maxRows);
		result.setData("rowfreeze", _rowFreeze);
		return result.toString();
	}
	
	/**
	 * Returns the maximum number of columns of this spreadsheet. you can assign
	 * new numbers by calling {@link #setMaxcolumns(int)}.<br/>
	 * If you do some operation on model will increase/decrease maxcolumn
	 * automatically, such as insert/delete columns into/from the sheet.
	 * 
	 * @return the maximum column number
	 */
	public int getMaxcolumns() {
		return _maxColumns;
	}

	/**
	 * Sets the maximum number of columns of this spread sheet. for example, if you
	 * set to 40, which means it allow column 0 to 39. the minimal value of
	 * max number of columns must large than 0;
	 * 
	 * @param maxcols  the maximum column
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

	private String getMaxcolumnsInJSON() {
		JSONObj result = new JSONObj();
		result.setData("maxcol", _maxColumns);
		result.setData("colfreeze", _colFreeze);
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
		Sheet sheet = getSelectedSheet();
		if (sheet != null) {
			PaneInformation pi = sheet.getPaneInformation();
			_rowFreeze = pi != null ? pi.getHorizontalSplitTopRow() - 1 : -1;
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
		if (_colFreeze > -1)
			return _colFreeze;
		Sheet sheet = getSelectedSheet();
		if (sheet != null) {
			PaneInformation pi = sheet.getPaneInformation();
			_colFreeze = pi != null ? pi.getVerticalSplitLeftColumn() - 1 : -1;
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
		Sheet sheet = getSelectedSheet();
		int rowHeight = sheet.getDefaultRowHeight();

		return (rowHeight <= 0) ? _defaultRowHeight : Utils.twipToPx(rowHeight);
	}

	
	/**
	 * Sets the default row height of the selected sheet
	 * @param rowHeight the row height
	 */
	public void setRowheight(int rowHeight) {
		Sheet sheet = getSelectedSheet();
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
		final Sheet sheet = getSelectedSheet();
		return Utils.getDefaultColumnWidthInPx(sheet);
	}

	/**
	 * Sets the default column width of the selected sheet
	 * @param columnWidth the default column width
	 */
	public void setColumnwidth(int columnWidth) {
		final Sheet sheet = getSelectedSheet();
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
		final Sheet sheet = getSelectedSheet();
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
	
	protected void renderProperties(ContentRenderer renderer) throws IOException {
		super.renderProperties(renderer);
		Sheet sheet = this.getSelectedSheet();

		renderer.render("rowHeight", getRowheight());
		renderer.render("columnWidth", getColumnwidth());

		int th = isHidecolumnhead() ? 1 : this.getTopheadheight();
		/**
		 * toph -> topPanelHeight
		 */
		renderer.render("topPanelHeight", th);
		int lw = isHiderowhead() ? 1 : this.getLeftheadwidth();
		renderer.render("leftPanelWidth", lw);

		/**
		 * cellpad -> cellPadding
		 */
		renderer.render("cellPadding", _cellpadding);
		renderer.render("scss", getDynamicMediaURI(this, _cssVersion++, "ss_" + this.getUuid(), "css"));

		renderer.render("maxRow", getMaxrowInJSON());
		renderer.render("maxColumn", getMaxcolumnsInJSON());

		renderer.render("sheetId", getSelectedSheetId());
		/**
		 * fs -> focusRect
		 */
		renderer.render("focusRect", getRectStr(_focusRect));
		/**
		 * sel -> selectionRect
		 */
		renderer.render("selectionRect", getRectStr(_selectionRect));
		/**
		 * hl -> highLightRect
		 */
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
		int rowBegin = getRowBegin(this);
		int rowEnd = Math.max(getRowEnd(this), getRowfreeze()) ;
		int colBegin = getColBegin(this);
		int colEnd = Math.max(getColEnd(this), getColumnfreeze());
		renderer.render("rowBegin", rowBegin);
		renderer.render("rowEnd", rowEnd);

		List<String> rowOutter = new ArrayList<String>();
		for (int i = rowBegin; i <= rowEnd; i++)
			rowOutter.add(getRowOuterAttrs(this, i));
		renderer.render("rowOuter", rowOutter);

		renderer.render("colBegin", colBegin);
		renderer.render("colEnd", colEnd);

		List<List> cellOuter = new ArrayList<List>();
		List<List> cellInner = new ArrayList<List>();
		List<List> celltext = new ArrayList<List>();
		for (int r = rowBegin; r <= rowEnd; r++) {
			List<String> outter = new ArrayList<String>();
			cellOuter.add(outter);

			List<String> inner = new ArrayList<String>();
			cellInner.add(inner);

			List<String> text = new ArrayList<String>();
			celltext.add(text);

			for (int c = colBegin; c <= colEnd; c++) {
				outter.add(getCellOuterAttrs(this, r, c));
				inner.add(getCellInnerAttrs(this, r, c));
				text.add(getCelltext(this, r, c));
			}
		}

		renderer.render("cellOuter", cellOuter);
		renderer.render("cellInner", cellInner);
		renderer.render("celltext", celltext);

		/**
		 * rename hidecolhead -> columnHeadHidden
		 */
		renderer.render("columnHeadHidden", _hideColhead);
		/**
		 * rename hiderowhead -> rowHeadHidden
		 */
		renderer.render("rowHeadHidden", _hideRowhead);

		if (!_hideColhead) {
			ArrayList<String> topHeaderOuter = new ArrayList<String>();
			ArrayList<String> topHeaderInner = new ArrayList<String>();
			ArrayList<String> columntitle = new ArrayList<String>();
			ArrayList<Boolean> topHeaderHiddens = new ArrayList<Boolean>();
			for (int i = colBegin; i <= colEnd; i++) {
				topHeaderOuter.add(getTopHeaderOuterAttrs(this, i));
				topHeaderInner.add(getTopHeaderInnerAttrs(this, i));
				columntitle.add(UtilFns.getColumntitle(this, i));
				topHeaderHiddens.add(UtilFns.getTopHeaderHiddens(this, i));
			}
			renderer.render("topHeaderOuter", topHeaderOuter.toArray());
			renderer.render("topHeaderInner", topHeaderInner.toArray());
			renderer.render("columntitle", columntitle.toArray());
			renderer.render("topHeaderHiddens", topHeaderHiddens.toArray());
		}
		if (!_hideRowhead) {
			ArrayList<String> leftHeaderOuter = new ArrayList<String>();
			ArrayList<String> leftHeaderInner = new ArrayList<String>();
			ArrayList<String> rowtitle = new ArrayList<String>();
			ArrayList<Boolean> leftHeaderHiddens = new ArrayList<Boolean>();
			for (int i = rowBegin; i <= rowEnd; i++) {
				leftHeaderOuter.add(UtilFns.getLeftHeaderOuterAttrs(this, i));
				leftHeaderInner.add(UtilFns.getLeftHeaderInnerAttrs(this, i));
				rowtitle.add(UtilFns.getRowtitle(this, i));
				leftHeaderHiddens.add(UtilFns.getLeftHeaderHiddens(this, i));
			}
			renderer.render("leftHeaderOuter", leftHeaderOuter.toArray());
			renderer.render("leftHeaderInner", leftHeaderInner.toArray());
			renderer.render("rowtitle", rowtitle.toArray());
			renderer.render("leftHeaderHiddens", leftHeaderHiddens.toArray());
		}
		renderer.render("dataPanel", ((SpreadsheetCtrl) this.getExtraCtrl()).getDataPanelAttrs());
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
	 * Return current selection rectangle. the returned value is a clone copy of
	 * current selection status. Default Value:(0,0,0,0)
	 * 
	 * @return current selection
	 */
	public Rect getSelection() {
		return (Rect) _selectionRect.clone();
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
			if (sel.getLeft() < 0 || sel.getTop() < 0
					|| sel.getRight() >= this.getMaxcolumns()
					|| sel.getBottom() >= this.getMaxrows()
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
		return (Rect) _highlightRect.clone();
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
				if (highlight.getLeft() < 0 || highlight.getTop() < 0
						|| highlight.getRight() >= this.getMaxcolumns()
						|| highlight.getBottom() >= this.getMaxrows()
						|| highlight.getLeft() > highlight.getRight()
						|| highlight.getTop() > highlight.getBottom()) {
					throw new UiException("illegal highlight : " + highlight.toString());
				}
				_highlightRect = (Rect) highlight.clone();
				result.setData("type", "show");
				result.setData("left", _highlightRect.getLeft());
				result.setData("top", _highlightRect.getTop());
				result.setData("right", _highlightRect.getRight());
				result.setData("bottom", _highlightRect.getBottom());
			}
			response("selectionHighlight", new AuHighlight(this, result.toString()));
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
				result = Spreadsheet.this.getAttribute(name, true);
			}

			return result;
		}
	}

	private class InnerFunctionMapper implements FunctionMapper, Serializable {
		private static final long serialVersionUID = 1L;

		public Collection getClassNames() {
			Page page = getPage();
			if (page != null) {
				return page.getFunctionMapper().getClassNames();
			}
			return null;
		}

		public Class resolveClass(String name) throws XelException {
			Page page = getPage();
			if (page != null) {
				return page.getFunctionMapper().resolveClass(name);
			}
			return null;
		}

		public Function resolveFunction(String prefix, String name)
				throws XelException {
			Page page = getPage();
			if (page != null) {
				return page.getFunctionMapper().resolveFunction(prefix, name);
			}
			return null;
		}

	}

	/* DataListener to handle sheet data event */
	private class InnerDataListener extends EventDispatchListener implements Serializable {
		private static final long serialVersionUID = 20100330164021L;

		public InnerDataListener() {
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
		}
		private Sheet getSheet(Ref rng) {
			return Utils.getSheetByRefSheet(_book, rng.getOwnerSheet()); 
		}
		private void onContentChange(SSDataEvent event) {
			final Ref rng = event.getRef();
			final Sheet sheet = getSheet(rng);
			if (!getSelectedSheet().equals(sheet))
				return;
			updateCell(sheet, rng.getLeftCol(), rng.getTopRow(), rng.getRightCol(), rng.getBottomRow());
		}
		private void onRangeInsert(SSDataEvent event) {
			final Ref rng = event.getRef();
			final Sheet sheet = getSheet(rng);
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
			final Sheet sheet = getSheet(rng);
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
			final Sheet sheet = getSheet(orng);
			if (!getSelectedSheet().equals(sheet))
				return;
			((ExtraCtrl) getExtraCtrl()).updateMergeCell(sheet, 
					rng.getLeftCol(), rng.getTopRow(), rng.getRightCol(), rng.getBottomRow(),
					orng.getLeftCol(), orng.getTopRow(), orng.getRightCol(), orng.getBottomRow());
		}
		private void onMergeAdd(SSDataEvent event) {
			final Ref rng = event.getRef();
			final Sheet sheet = getSheet(rng);
			if (!getSelectedSheet().equals(sheet))
				return;
			((ExtraCtrl) getExtraCtrl()).addMergeCell(sheet, 
					rng.getLeftCol(), rng.getTopRow(), rng.getRightCol(), rng.getBottomRow());

		}
		private void onMergeDelete(SSDataEvent event) {
			final Ref orng = event.getRef();
			final Sheet sheet = getSheet(orng);
			if (!getSelectedSheet().equals(sheet))
				return;
			((ExtraCtrl) getExtraCtrl()).deleteMergeCell(sheet, orng.getLeftCol(),
					orng.getTopRow(), orng.getRightCol(), orng.getBottomRow());
		}
		private void onSizeChange(SSDataEvent event) {
			//TODO shall pass the range over to the client side and let client side do it; rather than iterate each column and send multiple command
			final Ref rng = event.getRef();
			final Sheet sheet = getSheet(rng);
			if (!getSelectedSheet().equals(sheet))
				return;
			if (rng.isWholeColumn()) {
				final int left = rng.getLeftCol();
				final int right = rng.getRightCol();
				for (int c = left; c <= right; ++c) {
					updateColWidth(sheet, c);
				}
			} else if (rng.isWholeRow()) {
				final int top = rng.getTopRow();
				final int bottom = rng.getBottomRow();
				for (int r = top; r <= bottom; ++r) {
					updateRowHeight(sheet, r);
				}
			}
		}
	}
	
	private void updateColWidth(Sheet sheet, int col) {
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

	private void updateRowHeight(Sheet sheet, int row) {
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
	
	public MergeMatrixHelper getMergeMatrixHelper(Sheet sheet) {
		if (sheet != getSelectedSheet())
			throw new UiException("not current selected sheet ");
		String sheetName = sheet.getSheetName();
		MergeMatrixHelper mmhelper = (MergeMatrixHelper) getAttribute(sheetName+MERGE_MATRIX_KEY);
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
			setAttribute(sheetName+MERGE_MATRIX_KEY, mmhelper);
		} else {
			mmhelper.update(fzr, fzc);
		}
		return mmhelper;
	}
	
	private HeaderPositionHelper getRowPositionHelper(Sheet sheet) {
		return getPositionHelpers(sheet)[0];
	}
	
	private HeaderPositionHelper getColumnPositionHelper(Sheet sheet) {
		return getPositionHelpers(sheet)[1];
	}
	
	//[0] row position, [1] column position
	private HeaderPositionHelper[] getPositionHelpers(Sheet sheet) {
		if (sheet != getSelectedSheet())
			throw new UiException("not current selected sheet ");
		HeaderPositionHelper helper = (HeaderPositionHelper) getAttribute(ROW_SIZE_HELPER_KEY);

		int maxcol = 0;
		if (helper == null) {
			int defaultSize = this.getRowheight();
			
			List<HeaderPositionInfo> infos = new ArrayList<HeaderPositionInfo>();
			
			for(Row row: sheet) {
				final boolean hidden = row.getZeroHeight();
				final int height = Utils.getRowHeightInPx(row);
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

	private void updateCell(Sheet sheet, int left, int top, int right, int bottom) {
		if (this.isInvalidated())
			return;// since it is invalidate, we don't need to do anymore

		int row, col;
		String sheetId = Utils.getSheetId(sheet);
		if (!sheetId.equals(this.getSelectedSheetId()))
			return;
		left = left > 0 ? left - 1 : 0;// for border, when update a range, we
		// should also update the left,top +1 part
		top = top > 0 ? top - 1 : 0;
		for (int i = left; i <= right; i++) {
			for (int j = top; j <= bottom; j++) {
				row = j;
				col = i;

				// update cell only in block or in freeze panels
				if (row > _loadedRect.getBottom()
						|| (row < _loadedRect.getTop() && row > getRowfreeze()))
					continue;
				if (col > _loadedRect.getRight()
						|| (col < _loadedRect.getLeft() && row > getColumnfreeze()))
					continue;

				final Cell cell = Utils.getCell(sheet, row, col);
				// if(cell==null) continue;

				JSONObj result = new JSONObj();
				result.setData("r", row);
				result.setData("c", col);
				result.setData("type", "udcell");

				CellStyle style = (cell == null) ? null : cell.getCellStyle();
				boolean wrap = false;

				CellFormatHelper cfh = new CellFormatHelper(sheet, row, col, getMergeMatrixHelper(sheet));
				String st = cfh.getHtmlStyle();
				String ist = cfh.getInnerHtmlStyle();
				if (st != null && !"".equals(st))
					result.setData("st", st);// style of text cell.

				if (ist != null && !"".equals(ist))
					result.setData("ist", ist);// inner style of text cell

				if (style != null && style.getWrapText()) {
					wrap = true;
					result.setData("wrap", true);
				}
				if (cfh.hasRightBorder())
					result.setData("rbo", true);

				int textHAlign = (cell == null) ? -1 : BookHelper.getRealAlignment(cell);
				switch(textHAlign) {
				case CellStyle.ALIGN_CENTER:
				case CellStyle.ALIGN_CENTER_SELECTION:
					result.setData("hal", "c");
					break;
				case CellStyle.ALIGN_RIGHT:
					result.setData("hal", "r");
					break;
				}
				Hyperlink hlink = cell == null ? null : Utils.getHyperlink(cell);
				final FormatText ft = (cell == null) ? null : Utils.getFormatText(cell);
				
				final RichTextString rstr = ft != null && ft.isRichTextString() ? ft.getRichTextString() : null; 
				String text = rstr != null ? Utils.formatRichTextString(sheet, rstr, wrap) : ft != null ? Utils.escapeCellText(ft.getCellFormatResult().text, wrap, wrap) : "";
				if (hlink != null) {
					text = Utils.formatHyperlink(sheet, hlink, text, wrap);
				}
				result.setData("val", text);
				// responseUpdateCell(row + "_" + col + "_" + _updateCellId.last(), "", sheetId, result.toString());
				response(row + "_" + col + "_" + _updateCellId.last(), new AuDataUpdate(this, "", sheetId, result.toString()));
			}
		}
	}
	
	private HeaderPositionHelper myGetColumnPositionHelper(Sheet sheet, int maxcol) {
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
			Sheet sheet;
			if (getSelectedSheetId().equals(sheetId)) {
				sheet = getSelectedSheet();
			} else {
				sheet = Utils.getSheetById(_book, sheetId);
			}
			// update helper size first before sheet.setColumnWidth, or it will fire a SSDataEvent
			HeaderPositionHelper helper = Spreadsheet.this.getColumnPositionHelper(sheet);
			helper.setInfoValues(column, newsize, id, hidden);

			sheet.setColumnWidth(column, Utils.pxToFileChar256(newsize, ((Book)sheet.getWorkbook()).getDefaultCharWidth()));
		}

		public void setRowSize(String sheetId, int rownum, int newsize, int id, boolean hidden) {
			Sheet sheet;
			if (getSelectedSheetId().equals(sheetId)) {
				sheet = getSelectedSheet();
			} else {
				sheet = Utils.getSheetById(_book, sheetId);
			}
			Row row = sheet.getRow(rownum);
			if (row == null) {
				row = sheet.createRow(rownum);
			}
			row.setHeight((short)Utils.pxToTwip(newsize));
			HeaderPositionHelper helper = Spreadsheet.this.getRowPositionHelper(sheet);
			helper.setInfoValues(rownum, newsize, id, hidden);
		}

		public HeaderPositionHelper getColumnPositionHelper(String sheetId) {
			Sheet sheet;
			if (getSelectedSheetId().equals(sheetId)) {
				sheet = getSelectedSheet();
			} else {
				sheet = Utils.getSheetById(_book, sheetId);
			}
			HeaderPositionHelper helper = Spreadsheet.this
					.getColumnPositionHelper(sheet);
			return helper;
		}

		public HeaderPositionHelper getRowPositionHelper(String sheetId) {
			Sheet sheet;
			if (getSelectedSheetId().equals(sheetId)) {
				sheet = getSelectedSheet();
			} else {
				sheet = Utils.getSheetById(_book, sheetId);
			}
			HeaderPositionHelper helper = Spreadsheet.this
					.getRowPositionHelper(sheet);
			return helper;
		}

		public MergeMatrixHelper getMergeMatrixHelper(Sheet sheet) {
			return Spreadsheet.this.getMergeMatrixHelper(sheet);
		}

		public Rect getSelectionRect() {
			return (Rect) _selectionRect.clone();
		}

		public Rect getFocusRect() {
			return (Rect) _focusRect.clone();
		}

		public void setSelectionRect(int left, int top, int right, int bottom) {
			_selectionRect.set(left, top, right, bottom);
		}

		public void setFocusRect(int left, int top, int right, int bottom) {
			_focusRect.set(left, top, right, bottom);
		}

		public Rect getLoadedRect() {
			return (Rect) _loadedRect.clone();
		}

		public void setLoadedRect(int left, int top, int right, int bottom) {
			_loadedRect.set(left, top, right, bottom);
			getWidgetHandler().onLoadOnDemand(getSelectedSheet(), left, top, right, bottom);
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

		public String getRowOuterAttrs(int row) {
			Sheet sheet = getSelectedSheet();
			HeaderPositionHelper helper = Spreadsheet.this.getRowPositionHelper(sheet);
			StringBuffer sb = new StringBuffer();
			sb.append("class=\"zsrow");
			HeaderPositionInfo info = helper.getInfo(row);
			int zsh = -1;
			if (info != null) {
				zsh = info.id;
				sb.append(" zsh").append(zsh);
			}
			sb.append("\"");

			HTMLs.appendAttribute(sb, "z.r", row);

			if (zsh >= 0) {
				HTMLs.appendAttribute(sb, "z.zsh", zsh);
			}

			return sb.toString();
		}

		private StringBuffer appendMergeSClass(StringBuffer sb, int row, int col) {
			Sheet sheet = getSelectedSheet();
			MergeMatrixHelper mmhelper = getMergeMatrixHelper(sheet);
			// Cell cell = sheet.getCell(row,col);
			// if(cell!=null){
			boolean islefttop = mmhelper.isMergeRangeLeftTop(row, col);
			MergedRect block;

			if (islefttop) {
				block = mmhelper.getMergeRange(row, col);
				sb.append(" zsmerge").append(block.getId());
			} else if ((block = mmhelper.getMergeRange(row, col)) != null) {
				sb.append(" zsmergee");
			} else {
				// System.out.println("3>>>>>"+row+","+col);
			}

			// }
			return sb;
		}

		public String getCellOuterAttrs(int row, int col) {
			StringBuffer sb = new StringBuffer();
			Sheet sheet = getSelectedSheet();
			MergeMatrixHelper matrix = getMergeMatrixHelper(sheet);
			HeaderPositionHelper rowHelper = Spreadsheet.this.getRowPositionHelper(sheet);
			HeaderPositionHelper colHelper = Spreadsheet.this.getColumnPositionHelper(sheet);
			Cell cell = Utils.getCell(sheet, row, col);

			// class="zscell zscw${cstatus.index} zsrhi${rstatus.index} ${s:getCellSClass(self,rstatus.index,cstatus.index)}"
			sb.append("class=\"zscell");
			int zsh = -1;
			int zsw = -1;
			HeaderPositionInfo info = colHelper.getInfo(col);
			if (info != null) {
				zsw = info.id;
				sb.append(" zsw").append(zsw);
			}
			info = rowHelper.getInfo(row);
			if (info != null) {
				zsh = info.id;
				sb.append(" zshi").append(zsh);
			}
			appendMergeSClass(sb, row, col);
			sb.append("\" ");

			CellFormatHelper cfh = new CellFormatHelper(sheet, row, col,
					getMergeMatrixHelper(sheet));
			HTMLs.appendAttribute(sb, "style", cfh.getHtmlStyle());

			HTMLs.appendAttribute(sb, "z.r", row);
			HTMLs.appendAttribute(sb, "z.c", col);
			if (zsw >= 0) {
				HTMLs.appendAttribute(sb, "z.zsw", zsw);
			}
			if (zsh >= 0) {
				HTMLs.appendAttribute(sb, "z.zsh", zsh);
			}

			if (cell != null) {
				CellStyle style = cell.getCellStyle();
				
				if (style != null && style.getWrapText()) {
					HTMLs.appendAttribute(sb, "z.wrap", "t");
				}

				int textHAlign = BookHelper.getRealAlignment(cell);
				switch(textHAlign) {
				case CellStyle.ALIGN_CENTER:
				case CellStyle.ALIGN_CENTER_SELECTION:
					HTMLs.appendAttribute(sb, "z.hal", "c");
					break;
				case CellStyle.ALIGN_RIGHT:
					HTMLs.appendAttribute(sb, "z.hal", "r");
					break;
				}

				if (cfh.hasRightBorder()) {
					HTMLs.appendAttribute(sb, "z.rbo", "t");
				}
			}

			MergedRect block;
			if ((block = matrix.getMergeRange(row, col)) != null) {
				HTMLs.appendAttribute(sb, "z.merr", block.getRight());
				HTMLs.appendAttribute(sb, "z.merid", block.getId());
				HTMLs.appendAttribute(sb, "z.merl", block.getLeft());
			}

			return sb.toString();
		}

		public String getCellInnerAttrs(int row, int col) {
			Sheet sheet = getSelectedSheet();
			Cell cell = Utils.getCell(sheet, row, col);
			StringBuffer sb = new StringBuffer();
			HeaderPositionHelper rowHelper = Spreadsheet.this
					.getRowPositionHelper(sheet);
			HeaderPositionHelper colHelper = Spreadsheet.this
					.getColumnPositionHelper(sheet);

			// class="zscelltxt zscwi${cstatus.index} zsrhi${rstatus.index}"
			sb.append("class=\"zscelltxt");
			HeaderPositionInfo info = colHelper.getInfo(col);
			if (info != null) {
				sb.append(" zswi").append(info.id);
			}
			info = rowHelper.getInfo(row);
			if (info != null) {
				sb.append(" zshi").append(info.id);
			}
			sb.append("\" ");

			CellFormatHelper cfh = new CellFormatHelper(sheet, row, col,
					getMergeMatrixHelper(sheet));
			HTMLs.appendAttribute(sb, "style", cfh.getInnerHtmlStyle());

			if (cell != null) {
				// no thing to do now.
			}
			return sb.toString();
		}

		public String getTopHeaderOuterAttrs(int col) {
			Sheet sheet = getSelectedSheet();
			HeaderPositionHelper colHelper = Spreadsheet.this
					.getColumnPositionHelper(sheet);
			StringBuffer sb = new StringBuffer();

			// class="zstopcell zscw${status.index}" z.c="${status.index}"
			sb.append("class=\"zstopcell");
			int zsw = -1;
			HeaderPositionInfo info = colHelper.getInfo(col);
			if (info != null) {
				zsw = info.id;
				sb.append(" zsw").append(zsw);
			}
			sb.append("\" ");

			HTMLs.appendAttribute(sb, "z.c", col);
			if (zsw >= 0) {
				HTMLs.appendAttribute(sb, "z.zsw", zsw);
			}

			return sb.toString();
		}

		public String getTopHeaderInnerAttrs(int col) {
			Sheet sheet = getSelectedSheet();
			HeaderPositionHelper colHelper = Spreadsheet.this
					.getColumnPositionHelper(sheet);
			StringBuffer sb = new StringBuffer();

			// class="zstopcelltxt zscw${status.index}"
			sb.append("class=\"zstopcelltxt");
			HeaderPositionInfo info = colHelper.getInfo(col);
			if (info != null) {
				sb.append(" zswi").append(info.id);
			}
			sb.append("\" ");

			return sb.toString();
		}

		public String getLeftHeaderOuterAttrs(int row) {
			Sheet sheet = getSelectedSheet();
			HeaderPositionHelper rowHelper = Spreadsheet.this
					.getRowPositionHelper(sheet);
			StringBuffer sb = new StringBuffer();

			// class="zsleftcell zsrow zslrh${status.index}"
			// z.r="${status.index}"
			sb.append("class=\"zsleftcell zsrow");
			int zsh = -1;
			HeaderPositionInfo info = rowHelper.getInfo(row);
			if (info != null) {
				zsh = info.id;
				sb.append(" zslh").append(zsh);
			}
			sb.append("\" ");

			HTMLs.appendAttribute(sb, "z.r", row);
			if (zsh >= 0) {
				HTMLs.appendAttribute(sb, "z.zsh", zsh);
			}

			return sb.toString();
		}

		public String getLeftHeaderInnerAttrs(int row) {
			Sheet sheet = getSelectedSheet();
			HeaderPositionHelper rowHelper = Spreadsheet.this
					.getRowPositionHelper(sheet);
			StringBuffer sb = new StringBuffer();

			sb.append("class=\"zsleftcelltxt ");
			HeaderPositionInfo info = rowHelper.getInfo(row);
			if (info != null) {
				sb.append(" zslh").append(info.id);
			}
			sb.append("\" ");

			return sb.toString();
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

		public void insertColumns(Sheet sheet, int col, int size) {

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

			JSONObj result = new JSONObj();
			result.setData("type", "column");
			result.setData("col", col);
			result.setData("size", size);

			List extnm = new ArrayList();
			int right = size + _loadedRect.getRight();
			for (int i = col; i <= right; i++) {
				extnm.add(getColumntitle(i));
			}
			result.setData("extnm", extnm);

			HeaderPositionHelper colHelper = Spreadsheet.this
					.getColumnPositionHelper(sheet);

			colHelper.shiftMeta(col, size);
			_maxColumns += size;
			int cf = getColumnfreeze();
			if (cf >= col) {
				_colFreeze += size;
			}

			result.setData("maxcol", _maxColumns);
			result.setData("colfreeze", _colFreeze);

			/**
			 * insertrc_ -> insertRowColumn
			 */
			// smartUpdateValues("insertrc_"+Utils.nextUpdateId(),new Object[]{"",Utils.getId(sheet),result.toString()});
			response("insertRowColumn" + Utils.nextUpdateId(), new AuInsertRowColumn(Spreadsheet.this, "", Utils.getSheetId(sheet), result.toString()));

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

		private void updateRowHeights(Sheet sheet, int row, int n) {
			for(int r = 0; r < n; ++r) {
				updateRowHeight(sheet, r+row);
			}
		}
		
		private void updateColWidths(Sheet sheet, int col, int n) {
			for(int r = 0; r < n; ++r) {
				updateColWidth(sheet, r+col);
			}
		}
		public void insertRows(Sheet sheet, int row, int size) {
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

			JSONObj result = new JSONObj();
			result.setData("type", "row");
			result.setData("row", row);
			result.setData("size", size);

			List extnm = new ArrayList();
			int bottom = size + _loadedRect.getBottom();
			for (int i = row; i <= bottom; i++) {
				extnm.add(getRowtitle(i));
			}
			result.setData("extnm", extnm);

			HeaderPositionHelper rowHelper = Spreadsheet.this.getRowPositionHelper(sheet);
			rowHelper.shiftMeta(row, size);
			_maxRows += size;
			int rf = getRowfreeze();
			if (rf >= row) {
				_rowFreeze += size;
			}

			result.setData("maxrow", _maxRows);
			result.setData("rowfreeze", _rowFreeze);

			/**
			 * insertrc_ -> insertRowColumn.
			 */
			// smartUpdateValues("insertrc_"+Utils.nextUpdateId(),new Object[]{"",Utils.getId(sheet),result.toString()});
			response("insertRowColumn" + Utils.nextUpdateId(), new AuInsertRowColumn(Spreadsheet.this, "", Utils.getSheetId(sheet), result.toString()));

			_loadedRect.setBottom(bottom);

			// update surround cell
			int top = row;
			bottom = bottom + size - 1;
			bottom = bottom >= _maxRows - 1 ? _maxRows - 1 : bottom;
			updateCell(sheet, _loadedRect.getLeft(), top, _loadedRect.getRight(), bottom);
			
			// update the inserted row height
			updateRowHeights(sheet, row, size); //update row height
		}

		public void removeColumns(Sheet sheet, int col, int size) {
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

			HeaderPositionHelper colHelper = Spreadsheet.this
					.getColumnPositionHelper(sheet);
			colHelper.unshiftMeta(col, size);

			_maxColumns -= size;
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

			/**
			 * removerc_ -> removeRowColumn.
			 */
			// smartUpdateValues("removerc_"+Utils.nextUpdateId(),new Object[]{"",Utils.getId(sheet),result.toString()});
			response("removeRowColumn" + Utils.nextUpdateId(), new AuRemoveRowColumn(Spreadsheet.this, "", Utils.getSheetId(sheet), result.toString()));
			_loadedRect.setRight(right);

			// update surround cell
			int left = col;
			right = left;
			updateCell(sheet, left, _loadedRect.getTop(), right, _loadedRect.getBottom());
		}

		public void removeRows(Sheet sheet, int row, int size) {
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

			JSONObj result = new JSONObj();
			result.setData("type", "row");
			result.setData("row", row);
			result.setData("size", size);

			List extnm = new ArrayList();

			int bottom = _loadedRect.getBottom() - size;
			if (bottom < row) {
				bottom = row - 1;
			}
			for (int i = row; i <= bottom; i++) {
				extnm.add(getRowtitle(i));
			}
			result.setData("extnm", extnm);

			HeaderPositionHelper rowHelper = Spreadsheet.this.getRowPositionHelper(sheet);
			rowHelper.unshiftMeta(row, size);

			_maxRows -= size;
			int rf = getRowfreeze();
			if (rf > -1 && row <= rf) {
				if (row + size > rf) {
					_rowFreeze = row - 1;
				} else {
					_rowFreeze -= size;
				}
			}

			result.setData("maxrow", _maxRows);
			result.setData("rowfreeze", _rowFreeze);
			/**
			 * removerc_ -> removeColumn Sam, need test, need
			 * Utils.nextUpdateId()???
			 */
			// smartUpdateValues("removerc_"+Utils.nextUpdateId(),new Object[]{"",Utils.getId(sheet),result.toString()});
			response("removeRowColumn" + Utils.nextUpdateId(), new AuRemoveRowColumn(Spreadsheet.this, "", Utils.getSheetId(sheet), result.toString()));
			_loadedRect.setBottom(bottom);

			// update surround cell
			int top = row;
			bottom = top;
			updateCell(sheet, _loadedRect.getLeft(), top, _loadedRect.getRight(), bottom);
		}

		private void removeAffectedMergeRange(Sheet sheet, int type, int index) {
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

		public void updateMergeCell(Sheet sheet, int left, int top, int right,
				int bottom, int oleft, int otop, int oright, int obottom) {
			deleteMergeCell(sheet, oleft, otop, oright, obottom);
			addMergeCell(sheet, left, top, right, bottom);
		}

		public void deleteMergeCell(Sheet sheet, int left, int top, int right, int bottom) {
			MergeMatrixHelper mmhelper = this.getMergeMatrixHelper(sheet);
			List torem = new ArrayList();
			mmhelper.deleteMergeRange(left, top, right, bottom, torem);
			for (Iterator iter = torem.iterator(); iter.hasNext();) {
				MergedRect rect = (MergedRect) iter.next();

				updateMergeCell0(sheet, rect, "remove");
			}
			updateCell(sheet, left > 0 ? left - 1 : 0, top > 1 ? top - 1 : 0, right + 1, bottom + 1);
		}

		private void updateMergeCell0(Sheet sheet, MergedRect block, String type) {
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
			int w = helper.getStartPixel(block.getRight() + 1) - helper.getStartPixel(block.getLeft());
			result.setData("width", w);

			/**
			 * merge_ -> mergeCell
			 */
			response("mergeCell" + Utils.nextUpdateId(), new AuMergeCell(Spreadsheet.this, "", Utils.getSheetId(sheet), result.toString()));
		}

		public void addMergeCell(Sheet sheet, int left, int top, int right,	int bottom) {
			MergeMatrixHelper mmhelper = this.getMergeMatrixHelper(sheet);

			List toadd = new ArrayList();
			List torem = new ArrayList();
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
			updateCell(sheet, left > 0 ? left - 1 : 0, top > 1 ? top - 1 : 0, right + 1, bottom + 1);

		}

		//in pixel
		public void setColumnWidth(Sheet sheet, int col, int width, int id, boolean hidden) {
			JSONObj result = new JSONObj();
			result.setData("type", "column");
			result.setData("column", col);
			result.setData("width", width);
			result.setData("id", id);
			result.setData("hidden", hidden);
			/**
			 * rename size_col -> columnSize
			 */
			smartUpdate("columnSize", (Object) new Object[] { "", Utils.getSheetId(sheet), result.toString() }, true);
		}

		//in pixels
		public void setRowHeight(Sheet sheet, int row, int height, int id, boolean hidden) {
			JSONObj result = new JSONObj();
			result.setData("type", "row");
			result.setData("row", row);
			result.setData("height", height);
			result.setData("id", id);
			result.setData("hidden", hidden);
			/**
			 * rename size_row -> rowSize
			 */
			smartUpdate("rowSize", (Object) new Object[] { "", Utils.getSheetId(sheet), result.toString() }, true);
		}

		@Override
		public Boolean getLeftHeaderHiddens(int row) {
			Sheet sheet = getSelectedSheet();
			HeaderPositionHelper rowHelper = Spreadsheet.this
					.getRowPositionHelper(sheet);
			List<Boolean> hiddens = new ArrayList<Boolean>();
			StringBuffer sb = new StringBuffer();
			HeaderPositionInfo info = rowHelper.getInfo(row);
			return info == null ? Boolean.FALSE : Boolean.valueOf(info.hidden);
		}

		@Override
		public Boolean getTopHeaderHiddens(int col) {
			Sheet sheet = getSelectedSheet();
			HeaderPositionHelper colHelper = Spreadsheet.this
					.getColumnPositionHelper(sheet);
			List<Boolean> hiddens = new ArrayList<Boolean>();
			StringBuffer sb = new StringBuffer();
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

		Sheet sheet = getSelectedSheet();

		HeaderPositionHelper colHelper = this.getColumnPositionHelper(sheet);
		HeaderPositionHelper rowHelper = this.getRowPositionHelper(sheet);
		MergeMatrixHelper mmhelper = this.getMergeMatrixHelper(sheet);

		boolean hiderow = isHiderowhead();
		boolean hidecol = isHidecolumnhead();

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
		sb.append(name).append(" .zsdata{\n");
		sb.append("padding-top:").append(th).append("px;\n");
		sb.append("padding-left:").append(lw).append("px;\n");
		sb.append("}\n");

		// zcss.setRule(name+" .zsrow","height",rh+"px",true,sid);
		sb.append(name).append(" .zsrow{\n");
		sb.append("height:").append(rh).append("px;\n");
		sb.append("}\n");

		// zcss.setRule(name+" .zscell",["padding","height","width","line-height"],["0px "+cp+"px 0px "+cp+"px",cellheight+"px",cellwidth+"px",lh+"px"],true,sid);
		sb.append(name).append(" .zscell{\n");
		sb.append("padding:").append("0px " + cp + "px 0px " + cp + "px;\n");
		sb.append("height:").append(cellheight).append("px;\n");
		sb.append("width:").append(cellwidth).append("px;\n");
		// sb.append("line-height:").append(lh).append("px;\n");
		sb.append("}\n");

		// zcss.setRule(name+" .zscelltxt",["width","height"],[celltextwidth+"px",cellheight+"px"],true,sid);
		sb.append(name).append(" .zscelltxt{\n");
		sb.append("width:").append(celltextwidth).append("px;\n");
		sb.append("height:").append(cellheight).append("px;\n");
		sb.append("}\n");

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

		sb.append(name).append(" .zstop{\n");
		sb.append("left:").append(lw).append("px;\n");
		sb.append("height:").append(fzr > -1 ? toph - 1 : toph).append("px;\n");
		// sb.append("line-height:").append(toph).append("px;\n");
		sb.append("}\n");

		sb.append(name).append(" .zstopi{\n");
		sb.append("height:").append(toph).append("px;\n");
		sb.append("}\n");

		sb.append(name).append(" .zstophead{\n");
		sb.append("height:").append(topheadh).append("px;\n");
		sb.append("}\n");

		sb.append(name).append(" .zscornertop{\n");
		sb.append("left:").append(lw).append("px;\n");
		sb.append("height:").append(cornertoph).append("px;\n");
		sb.append("}\n");

		// relative, so needn't set top position.
		/*
		 * sb.append(name).append(" .zstopblock{\n");
		 * sb.append("top:").append(topblocktop).append("px;\n");
		 * sb.append("}\n");
		 */

		// zcss.setRule(name+" .zstopcell",["padding","height","width","line-height"],["0px "+cp+"px 0px "+cp+"px",th+"px",cellwidth+"px",lh+"px"],true,sid);
		sb.append(name).append(" .zstopcell{\n");
		sb.append("padding:").append("0px " + cp + "px 0px " + cp + "px;\n");
		sb.append("height:").append(topheadh).append("px;\n");
		sb.append("width:").append(cellwidth).append("px;\n");
		sb.append("line-height:").append(topheadh).append("px;\n");
		sb.append("}\n");

		// zcss.setRule(name+" .zstopcelltxt","width", celltextwidth
		// +"px",true,sid);
		sb.append(name).append(" .zstopcelltxt{\n");
		sb.append("width:").append(celltextwidth).append("px;\n");
		sb.append("}\n");

		// zcss.setRule(name+" .zsleft",["top","width"],[th+"px",(lw-2)+"px"],true,sid);

		int leftw = lw - 1;
		int leftheadw = leftw;
		int leftblockleft = leftw;

		if (fzc > -1) {
			leftw = leftw + colHelper.getStartPixel(fzc + 1);
		}

		sb.append(name).append(" .zsleft{\n");
		sb.append("top:").append(th).append("px;\n");
		sb.append("width:").append(fzc > -1 ? leftw - 1 : leftw)
				.append("px;\n");
		sb.append("}\n");

		sb.append(name).append(" .zslefti{\n");
		sb.append("width:").append(leftw).append("px;\n");
		sb.append("}\n");

		sb.append(name).append(" .zslefthead{\n");
		sb.append("width:").append(leftheadw).append("px;\n");
		sb.append("}\n");

		sb.append(name).append(" .zsleftblock{\n");
		sb.append("left:").append(leftblockleft).append("px;\n");
		sb.append("}\n");

		sb.append(name).append(" .zscornerleft{\n");
		sb.append("top:").append(th).append("px;\n");
		sb.append("width:").append(leftheadw).append("px;\n");
		sb.append("}\n");

		// zcss.setRule(name+" .zsleftcell",["height","line-height"],[(rh-1)+"px",(rh)+"px"],true,sid);//for
		// middle the text, i use row leight instead of lh
		sb.append(name).append(" .zsleftcell{\n");
		sb.append("height:").append(rh - 1).append("px;\n");
		sb.append("line-height:").append(rh - 1).append("px;\n");
		sb.append("}\n");

		// zcss.setRule(name+" .zscorner",["width","height"],[(lw-2)+"px",(th-2)+"px"],true,sid);

		sb.append(name).append(" .zscorner{\n");
		sb.append("width:").append(fzc > -1 ? leftw : leftw + 1)
				.append("px;\n");
		sb.append("height:").append(fzr > -1 ? toph : toph + 1).append("px;\n");
		sb.append("}\n");

		sb.append(name).append(" .zscorneri{\n");
		sb.append("width:").append(lw - 2).append("px;\n");
		sb.append("height:").append(th - 2).append("px;\n");
		sb.append("}\n");

		sb.append(name).append(" .zscornerblock{\n");
		sb.append("left:").append(lw).append("px;\n");
		sb.append("top:").append(th).append("px;\n");
		sb.append("}\n");

		// zcss.setRule(name+" .zshboun","height",th+"px",true,sid);
		sb.append(name).append(" .zshboun{\n");
		sb.append("height:").append(th).append("px;\n");
		sb.append("}\n");

		// zcss.setRule(name+" .zshboun","height",th+"px",true,sid);
/*		sb.append(name).append(" .zshboun{\n");
		sb.append("height:").append(th).append("px;\n");
		sb.append("}\n");
*/
		// zcss.setRule(name+" .zshbouni","height",th+"px",true,sid);
		sb.append(name).append(" .zshbouni{\n");
		sb.append("height:").append(th).append("px;\n");
		sb.append("}\n");

		sb.append(name).append(" .zsfztop{\n");
		sb.append("border-bottom-style:").append(fzr > -1 ? "solid" : "none")
				.append(";");
		sb.append("}\n");
		sb.append(name).append(" .zsfzcorner{\n");
		sb.append("border-bottom-style:").append(fzr > -1 ? "solid" : "none")
				.append(";");
		sb.append("}\n");

		sb.append(name).append(" .zsfzleft{\n");
		sb.append("border-right-style:").append(fzc > -1 ? "solid" : "none")
				.append(";");
		sb.append("}\n");
		sb.append(name).append(" .zsfzcorner{\n");
		sb.append("border-right-style:").append(fzc > -1 ? "solid" : "none")
				.append(";");
		sb.append("}\n");

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
				sb
						.append("_filter:chroma(color=" + color_to_transparent
								+ ");");
			}
			sb.append("}\n");
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
				sb.append(name).append(" .zsw").append(cid).append("{\n");
				sb.append("display:none;\n");
				sb.append("}\n");

			} else {
				sb.append(name).append(" .zsw").append(cid).append("{\n");
				sb.append("width:").append(cellwidth).append("px;\n");
				sb.append("}\n");

				sb.append(name).append(" .zswi").append(cid).append("{\n");
				sb.append("width:").append(celltextwidth).append("px;\n");
				sb.append("}\n");
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
				sb.append("}\n");

				sb.append(name).append(" .zslh").append(cid).append("{\n");
				sb.append("display:none;");
				sb.append("}\n");

			} else {
				sb.append(name).append(" .zsh").append(cid).append("{\n");
				sb.append("height:").append(height).append("px;");
				sb.append("}\n");

				sb.append(name).append(" .zshi").append(cid).append("{\n");
				sb.append("height:").append(cellheight).append("px;");
				sb.append("}\n");

				int h2 = (height < 1) ? 0 : height - 1;

				sb.append(name).append(" .zslh").append(cid).append("{\n");
				sb.append("height:").append(h2).append("px;");
				sb.append("line-height:").append(h2).append("px;");
				sb.append("}\n");

			}
		}
		sb.append(".zs_header{}\n");// for indicating add new rule before this

		// merge size;
		List ranges = mmhelper.getRanges();
		Iterator iter = ranges.iterator();

		while (iter.hasNext()) {
			MergedRect block = (MergedRect) iter.next();
			int left = block.getLeft();
			int right = block.getRight();
			int width = 0;
			for (int i = left; i <= right; i++) {
				final HeaderPositionInfo info = colHelper.getInfo(i);
				final boolean hidden = info.hidden;
				final int colSize = hidden ? 0 : info.size;
				width += colSize;
			}

			if (width <= 0) { //total hidden
				sb.append(name).append(" .zsmerge").append(block.getId()).append("{\n");
				sb.append("display:none;");
				sb.append("}\n");

				sb.append(name).append(" .zsmerge").append(block.getId());
				sb.append(" .zscelltxt").append("{\n");
				sb.append("display:none;");
				sb.append("}\n");
			} else {
				celltextwidth = width - 2 * cp - 1;// 1 is border width
	
				if (!isGecko) {
					cellwidth = celltextwidth;
				} else {
					cellwidth = width;
				}
				sb.append(name).append(" .zsmerge").append(block.getId()).append("{\n");
				sb.append("width:").append(cellwidth).append("px;");
				sb.append("}\n");
	
				sb.append(name).append(" .zsmerge").append(block.getId());
				sb.append(" .zscelltxt").append("{\n");
				sb.append("width:").append(celltextwidth).append("px;");
				sb.append("}\n");
			}
		}

		sb.append(".zs_indicator{}\n");// for indicating the css is load ready

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
				sb.append(comp.getId());
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

	private void doSheetClean(Sheet sheet) {
		List list = loadWidgetLoaders();
		int size = list.size();
		for (int i = 0; i < size; i++) {
			((WidgetLoader) list.get(i)).onSheetClean(sheet);
		}
		removeAttribute(MERGE_MATRIX_KEY);
		_loadedRect.set(-1, -1, -1, -1);
		_selectionRect.set(0, 0, 0, 0);
		_focusRect.set(0, 0, 0, 0);
	}

	private void doSheetSelected(Sheet sheet) {
		List list = loadWidgetLoaders();
		int size = list.size();
		for (int i = 0; i < size; i++) {
			((WidgetLoader) list.get(i)).onSheetSelected(sheet);
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
		Sheet sheet = getSelectedSheet();

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
	/* package */List loadWidgetLoaders() {
		if (_widgetLoaders != null)
			return _widgetLoaders;
		_widgetLoaders = new ArrayList();
		String loaderclzs = (String) getAttribute(WIDGET_LOADERS_KEY);
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
		String handlerclz = (String) getAttribute(WIDGET_HANDLER_KEY);
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
	private class VoidWidgetHandler implements WidgetHandler {

		public boolean addWidget(Widget widget) {
			return false;
		}

		public Spreadsheet getSpreadsheet() {
			return Spreadsheet.this;
		};

		public void invaliate() {
		}

		public void onLoadOnDemand(Sheet sheet, int left, int top, int right, int bottom) {
		}

		public boolean removeWidget(Widget widget) {
			return false;
		}

		public void init(Spreadsheet spreadsheet) {
		}
	}

	private void processStartEditing(String token, StartEditingEvent event) {
		if (!event.isCancel()) {
			Object val;
			if (event.isEditingSet() || event.getClientValue() == null) {
				val = event.getEditingValue();
			} else {
				val = event.getClientValue();
			}

			processStartEditing0(token, event.getSheet(), event.getRow(), event
					.getColumn(), val);
		} else {
			processCancelEditing0(token, event.getSheet(), event.getRow(),
					event.getColumn());
		}
	}

	private void processStopEditing(String token, StopEditingEvent event) {
		if (!event.isCancel()) {
			processStopEditing0(token, event.getSheet(), event.getRow(), event.getColumn(), event.getEditingValue());
		} else
			processCancelEditing0(token, event.getSheet(), event.getRow(), event.getColumn());
	}

	private void processStopEditing0(String token, Sheet sheet, int row, int col, Object value) {
		try {
			Utils.setEditText(sheet, row, col, value == null ? "" : value.toString());

			JSONObj result = new JSONObj();
			result.setData("r", row);
			result.setData("c", col);
			result.setData("type", "stopedit");
			result.setData("val", "");

			// responseUpdateCell("stop", token, Utils.getId(sheet), result.toString());
			smartUpdate("dataUpdateStop", new String[] { token,	Utils.getSheetId(sheet), result.toString() });

		} catch (RuntimeException x) {
			processCancelEditing0(token, sheet, row, col);
			throw x;
		}
	}

	private void processStartEditing0(String token, Sheet sheet, int row, int col, Object value) {
		try {
			JSONObj result = new JSONObj();
			result.setData("r", row);
			result.setData("c", col);
			result.setData("type", "startedit");
			result.setData("val", value == null ? "" : value.toString());

			// responseUpdateCell("start", token, Utils.getId(sheet), result.toString());
			smartUpdate("dataUpdateStart", new String[] { token, Utils.getSheetId(sheet), result.toString() });
		} catch (RuntimeException x) {
			processCancelEditing0(token, sheet, row, col);
			throw x;
		}
	}

	private void processCancelEditing0(String token, Sheet sheet, int row, int col) {
		JSONObj result = new JSONObj();
		result.setData("r", row);
		result.setData("c", col);
		result.setData("type", "canceledit");
		result.setData("val", "");
		// responseUpdateCell("cancel", token, Utils.getId(sheet), result.toString());
		smartUpdate("dataUpdateCancel", new String[] { token, Utils.getSheetId(sheet), result.toString() });
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
	 * remove editor's focus on specified name
	 */
	public void removeEditorFocus(String name){
		//smartUpdateValues("removeEditorFocus",new Object[]{name});
		response("removeEditorFocus", new AuInvoke((Component)this,"removeEditorFocus", name));
		removeFocus(name);
	}
	
	/**
	 *  add and move other editor's focus
	 */
	public void moveEditorFocus(String name, String color, int row ,int col){
		//smartUpdateValues("moveEditorFocus",new Object[]{name ,color, ""+row ,""+col});
		//System.out.println("editor "+this.editor.getSender()+" recv: "+name+ " moveFocused");
		response("moveEditorFocus", new AuInvoke((Component)this,"moveEditorFocus", new String[]{name, color,""+row,""+col}));
		removeFocus(name);
		_focuses.add(new Focus(name, color, row, col));
	
	}
	/**
	 * remove editor focus on specified editor name
	 */
	private void removeFocus(String name){
		Focus targetFocus = null;
		Iterator iter = _focuses.iterator();
		for(;iter.hasNext();){
			Focus _focus=(Focus)iter.next();
			if((_focus).name.equals(name)){
				targetFocus=_focus;
				break;
			}
		}
		_focuses.remove(targetFocus);
	}
	
	
	/**
	 * Sam. added for zss app, *******Add focus function*************
	 */
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
		Iterator iter = _focuses.iterator();
		for(;iter.hasNext();){
			Focus _focus=(Focus)iter.next();
			row=_focus.row;
			col=_focus.col;
			if((col>=left && col<=right) ||
			   (row>=top  && row<=bottom)){
				this.moveEditorFocus(_focus.name, _focus.color, row, col);
			}
		}
	}
	
	/**
	 * Sam added for zss app
	 * @param sheet
	 */
	//it will be call when delete sheet 
/*	public void cleanRelatedState(Sheet sheet){
		stateManager.cleanRelatedState(sheet);
	}
*/

	static {
		addClientEvent(Spreadsheet.class, Events.ON_CELL_FOUCSED, CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, Events.ON_CELL_SELECTION,	CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, Events.ON_SELECTION_CHANGE, CE_IMPORTANT | CE_DUPLICATE_IGNORE | CE_NON_DEFERRABLE);
		addClientEvent(Spreadsheet.class, Events.ON_CELL_CLICK, CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, Events.ON_CELL_RIGHT_CLICK, CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, Events.ON_CELL_DOUBLE_CLICK, CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, Events.ON_HEADER_CLICK, CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, Events.ON_HEADER_RIGHT_CLICK,	CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, Events.ON_HEADER_DOUBLE_CLICK, CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, Events.ON_HYPERLINK, 0);
		addClientEvent(Spreadsheet.class, org.zkoss.zk.ui.event.Events.ON_CTRL_KEY, CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, org.zkoss.zk.ui.event.Events.ON_BLUR,	CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		
		addClientEvent(Spreadsheet.class, "onZSSStartEditing", CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, "onZSSEditboxEditing", CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, "onZSSStopEditing", CE_IMPORTANT | CE_NON_DEFERRABLE | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, "onZSSCellFocused", CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, "onZSSCellFetch", CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, "onZSSSyncBlock", CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, "onZSSHeaderModif", CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, "onZSSCellMouse", CE_IMPORTANT | CE_DUPLICATE_IGNORE);
		addClientEvent(Spreadsheet.class, "onZSSHeaderMouse", CE_IMPORTANT | CE_DUPLICATE_IGNORE);
	}

	final Command[] Commands = { new BlockSyncCommand(),
			new CellFetchCommand(), new CellSelectionCommand(),
			new CellFocusedCommand(), new CellMouseCommand(),
			new SelectionChangeCommand(), new HeaderMouseCommand(),
			new HeaderCommand(), new StartEditingCommand(),
			new StopEditingCommand(), new EditboxEditingCommand()};

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
}