/* RangeImpl.java

	Purpose:
		
	Description:
		
	History:
		Mar 10, 2010 2:54:55 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.impl;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.SortedMap;
import java.util.TreeMap;

import org.zkoss.poi.ss.SpreadsheetVersion;
import org.zkoss.poi.ss.usermodel.AutoFilter;
import org.zkoss.poi.ss.usermodel.BorderStyle;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.CellValue;
import org.zkoss.poi.ss.usermodel.Chart;
import org.zkoss.poi.ss.usermodel.ClientAnchor;
import org.zkoss.poi.ss.usermodel.DataValidation;
import org.zkoss.poi.ss.usermodel.FilterColumn;
import org.zkoss.poi.ss.usermodel.FormulaError;
import org.zkoss.poi.ss.usermodel.Hyperlink;
import org.zkoss.poi.ss.usermodel.Picture;
import org.zkoss.poi.ss.usermodel.RichTextString;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.poi.ss.usermodel.charts.ChartData;
import org.zkoss.poi.ss.usermodel.charts.ChartGrouping;
import org.zkoss.poi.ss.usermodel.charts.ChartType;
import org.zkoss.poi.ss.usermodel.charts.LegendPosition;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.poi.xssf.usermodel.XSSFChartX;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.EventQueue;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.engine.RefBook;
import org.zkoss.zss.engine.RefSheet;
import org.zkoss.zss.engine.event.SSDataEvent;
import org.zkoss.zss.engine.impl.AreaRefImpl;
import org.zkoss.zss.engine.impl.CellRefImpl;
import org.zkoss.zss.engine.impl.ChangeInfo;
import org.zkoss.zss.engine.impl.MergeChange;
import org.zkoss.zss.engine.impl.RefAddr;
import org.zkoss.zss.engine.impl.RefSheetImpl;
import org.zkoss.zss.model.Areas;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.FormatText;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zul.Messagebox;
/**
 * Implementation of {@link Range} which plays a facade to operate on the spreadsheet data models
 * and maintain the reference dependency. 
 * @author henrichen
 *
 */
public class RangeImpl implements Range {
	private final Worksheet _sheet;
	private int _left = Integer.MAX_VALUE;
	private int _top = Integer.MAX_VALUE;
	private int _right = Integer.MIN_VALUE;
	private int _bottom = Integer.MIN_VALUE;
	private Set<Ref> _refs;

	public RangeImpl(int row, int col, Worksheet sheet, Worksheet lastSheet) {
		_sheet = sheet;
		initCellRef(row, col, sheet, lastSheet);
	}
	
	private void initCellRef(int row, int col, Worksheet sheet, Worksheet lastSheet) {
		final Book book  = (Book) sheet.getWorkbook();
		final int s1 = book.getSheetIndex(sheet);
		final int s2 = book.getSheetIndex(lastSheet);
		final int sb = Math.min(s1,s2);
		final int se = Math.max(s1,s2);
		for (int s = sb; s <= se; ++s) {
			final Worksheet sht = book.getWorksheetAt(s);
			final RefSheet refSht = BookHelper.getOrCreateRefBook(book).getOrCreateRefSheet(sht.getSheetName());
			addRef(new CellRefImpl(row, col, refSht));
		}
	}
	
	public RangeImpl(int tRow, int lCol, int bRow, int rCol, Worksheet sheet, Worksheet lastSheet) {
		_sheet = sheet;
		if (tRow == bRow && lCol == rCol) 
			initCellRef(tRow, lCol, sheet, lastSheet);
		else
			initAreaRef(tRow, lCol, bRow, rCol, sheet, lastSheet);
	}
	
	/*package*/ RangeImpl(Ref ref, Worksheet sheet) {
		_sheet = BookHelper.getSheet(sheet, ref.getOwnerSheet());
		addRef(ref);
	}
	
	private RangeImpl(Set<Ref> refs, Worksheet sheet) {
		_sheet = sheet;
		for(Ref ref : refs) {
			addRef(ref);
		}
	}
	
	private void initAreaRef(int tRow, int lCol, int bRow, int rCol, Worksheet sheet, Worksheet lastSheet) {
		final Book book  = (Book) sheet.getWorkbook();
		final int s1 = book.getSheetIndex(sheet);
		final int s2 = book.getSheetIndex(lastSheet);
		final int sb = Math.min(s1,s2);
		final int se = Math.max(s1,s2);
		for (int s = sb; s <= se; ++s) {
			final Worksheet sht = book.getWorksheetAt(s);
			final RefSheet refSht = BookHelper.getOrCreateRefBook(book).getOrCreateRefSheet(sht.getSheetName());
			addRef(new AreaRefImpl(tRow, lCol, bRow, rCol, refSht));
		}
	}
	
	public Collection<Ref> getRefs() {
		if (_refs == null) {
			_refs = new LinkedHashSet<Ref>(3);
		}
		return _refs;
	}
	
	@Override
	public Worksheet getSheet() {
		return _sheet;
	}

	@Override
	public Hyperlink getHyperlink() {
		synchronized(_sheet) {
			Ref ref = _refs != null && !_refs.isEmpty() ? _refs.iterator().next() : null;
			if (ref != null) {
				final int tRow = ref.getTopRow();
				final int lCol = ref.getLeftCol();
				final RefSheet refSheet = ref.getOwnerSheet();
				final Cell cell = getCell(tRow, lCol, refSheet);
				if (cell != null)
					return BookHelper.getHyperlink(cell);
			}
			return null;
		}
	}
	@Override
	public void setHyperlink(int linkType, String address, String display) {
		synchronized(_sheet) {
			if (display == null) {
				display = address;
			}
			new HyperlinkSetter().setValue(new HyperlinkContext(linkType, address, display));
			final boolean old = isDirectHyperlink();
			try {
				setDirectHyperlink(true); //avoid setEditText() recursive to HyperlinkStringSetter#setCellValue()
				setEditText(display);
			} finally {
				setDirectHyperlink(old);
			}
		}
	}

	private int getHyperlinkType(String addr) {
		if (addr != null && !isDirectHyperlink()) {
			if (addr.toUpperCase().startsWith("HTTP://")) {
				return Hyperlink.LINK_URL;
			}
		}
		return -1;
	}
	
	private boolean _directHyperlink;
	
	private boolean isDirectHyperlink() {
		return _directHyperlink;
	}
	
	private void setDirectHyperlink(boolean b) {
		_directHyperlink = b;
	}
	
	private class HyperlinkSetter extends ValueSetter {
		protected Set<Ref>[] setCellValue(int row, int col, RefSheet refSheet, Object value) {
			final HyperlinkContext context = (HyperlinkContext)value;
			final Cell cell = BookHelper.getOrCreateCell(_sheet, row, col);
			BookHelper.setCellHyperlink(cell, context.getLinktype(), context.getAddress());
			return null; //not need to return Set<Ref>, setHyperlink will call setEditText() and it will do the update
		}
	}
	
	private class HyperlinkStringSetter extends ValueSetter {
		
		protected Set<Ref>[] setCellValue(int row, int col, RefSheet refSheet, Object value) {
			final HyperlinkContext context = (HyperlinkContext)value;
			final Cell cell = BookHelper.getOrCreateCell(_sheet, row, col);
			final Set<Ref>[] refs = BookHelper.setCellValue(cell, context.getDisplay());
			BookHelper.setCellHyperlink(cell, context.getLinktype(), context.getAddress());
			return refs;
		}
		
	}
	
	private static class HyperlinkContext {
		private int _linkType;
		private String _address;
		private String _display;
		
		HyperlinkContext(int linkType, String address, String display) {
			_linkType = linkType;
			_address = address;
			_display = display;
		}
		
		public int getLinktype() {
			return _linkType;
		}
		
		public String getAddress() {
			return _address;
		}
		
		public String getDisplay() {
			return _display;
		}
	}
	@Override
	public RichTextString getText() {
		synchronized(_sheet) {
			Ref ref = _refs != null && !_refs.isEmpty() ? _refs.iterator().next() : null;
			if (ref != null) {
				final int tRow = ref.getTopRow();
				final int lCol = ref.getLeftCol();
				final RefSheet refSheet = ref.getOwnerSheet();
				final Cell cell = getCell(tRow, lCol, refSheet);
				if (cell != null)
					return BookHelper.getText(cell);
			}
			return null;
		}
	}
	@Override
	public FormatText getFormatText() {
		synchronized(_sheet) {
			Ref ref = _refs != null && !_refs.isEmpty() ? _refs.iterator().next() : null;
			if (ref != null) {
				final int tRow = ref.getTopRow();
				final int lCol = ref.getLeftCol();
				final RefSheet refSheet = ref.getOwnerSheet();
				final Cell cell = getCell(tRow, lCol, refSheet);
				if (cell != null)
					return BookHelper.getFormatText(cell);
			}
			return null;
		}
	}
	
	@Override
	public RichTextString getRichEditText() {
		synchronized(_sheet) {
			Ref ref = _refs != null && !_refs.isEmpty() ? _refs.iterator().next() : null;
			if (ref != null) {
				final int tRow = ref.getTopRow();
				final int lCol = ref.getLeftCol();
				final RefSheet refSheet = ref.getOwnerSheet();
				final Cell cell = getCell(tRow, lCol, refSheet);
				if (cell != null) {
					return BookHelper.getRichEditText(cell);
				}
			}
			return null;
		}
	}
	
	@Override
	public void setRichEditText(RichTextString rstr) {
		synchronized(_sheet) {
			setValue(rstr);
		}
	}
	
	@Override
	public String getEditText() {
		synchronized(_sheet) {
			Ref ref = _refs != null && !_refs.isEmpty() ? _refs.iterator().next() : null;
			if (ref != null) {
				final int tRow = ref.getTopRow();
				final int lCol = ref.getLeftCol();
				final RefSheet refSheet = ref.getOwnerSheet();
				final Cell cell = getCell(tRow, lCol, refSheet);
				if (cell != null) {
					return BookHelper.getEditText(cell);
				}
			}
			return null;
		}
	}
	
	//return null if a valid input; otherwise the associated DataVailation for invalid input.
	@Override
	public DataValidation validate(String txt) {
		synchronized(_sheet) {
			Ref ref = _refs != null && !_refs.isEmpty() ? _refs.iterator().next() : null;
			if (ref != null) {
				final int tRow = ref.getTopRow();
				final int lCol = ref.getLeftCol();
				final RefSheet refSheet = ref.getOwnerSheet();
				final Cell cell = getCell(tRow, lCol, refSheet);
				
				final Object[] values = BookHelper.editTextToValue(txt, cell);
				final Worksheet sheet = BookHelper.getSheet(_sheet, refSheet);
				
				final int cellType = values == null ? -1 : ((Integer)values[0]).intValue();
				final Object value = values == null ? null : values[1];
				
				return BookHelper.validate(sheet, tRow, lCol, value, cellType);
			}
			return null;
		}
	}
	
	@Override
	public void setEditText(String txt) {
		synchronized(_sheet) {
			Ref ref = _refs != null && !_refs.isEmpty() ? _refs.iterator().next() : null;
			if (ref != null) {
				final int tRow = ref.getTopRow();
				final int lCol = ref.getLeftCol();
				final RefSheet refSheet = ref.getOwnerSheet();
				final Cell cell = getCell(tRow, lCol, refSheet);
				
				final Object[] values = BookHelper.editTextToValue(txt, cell);
				final Worksheet sheet = BookHelper.getSheet(_sheet, refSheet);
				
				final int cellType = values == null ? -1 : ((Integer)values[0]).intValue();
				final Object value = values == null ? null : values[1];
				
				Set<Ref>[] refs = null;
				if (cellType != -1) {
					switch(((Integer)values[0]).intValue()) {
					case Cell.CELL_TYPE_FORMULA:
						refs = setFormula((String)values[1]); //Formula
						break;
					case Cell.CELL_TYPE_STRING:
						refs = setValue((String)values[1]); //String
						break;
					case Cell.CELL_TYPE_BOOLEAN:
						refs = setValue((Boolean)values[1]); //boolean
						break;
					case Cell.CELL_TYPE_NUMERIC:
						final Object val = values[1];
						refs = val instanceof Number ?
							setValue((Number)val): //number
							setValue((Date)val); //date
						if (values.length > 2 && values[2] != null) {
							setDateFormat((String) values[2]);
						}
						break;
					case Cell.CELL_TYPE_ERROR:
						refs = setValue((Byte)values[1]);
						break;
					}
				} else {
					refs = setValue((String) null); 
				}
				reevaluateAndNotify(refs);
			}
		}
	}
	
	private void setDateFormat(String formatString) {
		for(Ref ref : _refs) {
			final int tRow = ref.getTopRow();
			final int lCol = ref.getLeftCol();
			final int bRow = ref.getBottomRow();
			final int rCol = ref.getRightCol();
			final RefSheet refSheet = ref.getOwnerSheet();
			final Worksheet sheet = BookHelper.getSheet(_sheet, refSheet);
			final Workbook book = sheet.getWorkbook();
			
			for(int row = tRow; row <= bRow; ++row) {
				for (int col = lCol; col <= rCol; ++col) {
					final Cell cell = BookHelper.getCell(sheet, row, col);
					final CellStyle style = cell != null ? cell.getCellStyle() : null;
					final String oldFormat = style != null ? style.getDataFormatString() : null;
					if (oldFormat == null || "General".equals(oldFormat)) {
						CellStyle newCellStyle = book.createCellStyle();
						if (style != null) {
							newCellStyle.cloneStyleFrom(style);
						}
						BookHelper.setDataFormat(book, newCellStyle, formatString); //prepare a DataFormat with the specified formatString
						BookHelper.setCellStyle(sheet, row, col, row, col, newCellStyle);
					}
				}
			}
		}
	}
	
	/*package*/ void reevaluateAndNotify(Set<Ref>[] refs) { //[0]: last, [1]: all
		if (refs != null) {
			final Book book  = (Book) _sheet.getWorkbook();
			BookHelper.reevaluateAndNotify(book, refs[0], refs[1]);
		}
	}
	
	/*package*/ Set<Ref>[] setValue(String value) {
		final int linkType = getHyperlinkType(value);
		return linkType > 0 ? new HyperlinkStringSetter().setValue(new HyperlinkContext(linkType, value, value)) : new StringValueSetter().setValue(value);
	}
	
	/*package*/ Set<Ref>[] setValue(Number value) {
		return new NumberValueSetter().setValue(value);
	}
	
	/*package*/ Set<Ref>[] setValue(Boolean value) {
		return new BooleanValueSetter().setValue(value);
	}
	
	/*package*/ Set<Ref>[] setValue(Date value) {
		return new DateValueSetter().setValue(value);
	}
	
	/*package*/ Set<Ref>[] setValue(RichTextString value) {
		return new RichTextStringValueSetter().setValue(value);
	}
	
	/*package*/ Set<Ref>[] setValue(Byte error) {
		return new ErrorValueSetter().setValue(error);
	}

	/*package*/ Set<Ref>[] setFormula(String value) {
		return new FormulaValueSetter().setValue(value);
	}
	
	private abstract class ValueSetter {
		private Set<Ref> _last;
		private Set<Ref> _all;
		private ValueSetter() {
			_last = new HashSet<Ref>();
			_all = new HashSet<Ref>();
		}
		/**
		 * Set specified value into cells specified by this Range.  
		 * @param value the value to be set into the cell
		 * @Return Ref set array: [0] is the last formula cells to be re-evaluated, [1] is the cells that has changed
		 */
		@SuppressWarnings("unchecked")
		public Set<Ref>[] setValue(Object value) {
			for(Ref ref : _refs) {
				final int tRow = ref.getTopRow();
				final int lCol = ref.getLeftCol();
				final int bRow = ref.getBottomRow();
				final int rCol = ref.getRightCol();
				final RefSheet refSheet = ref.getOwnerSheet();
				
				for(int row = tRow; row <= bRow; ++row) {
					for (int col = lCol; col <= rCol; ++col) {
						Set<Ref>[] refs = setCellValue(row, col, refSheet, value);
						if (refs != null) {
							_last.addAll(refs[0]);
							_all.addAll(refs[1]);
						}
					}
				}
				_all.add(ref); //this reference shall be reloaded, too.
			}
			return (Set<Ref>[]) new Set[] {_last, _all};
		}
		
		/**
		 * Set specified value into cell at the specified row and column.  
		 * @param rowIndex row index
		 * @param colIndex column index
		 * @param value the value to be set into the cell
		 * @Return Ref set arrary: [0] are the cells to be re-evaluated, [1] are the cells that were consequently updated 
		 */
		protected abstract Set<Ref>[] setCellValue(int row, int col, RefSheet refSheet, Object value);
	}

	private class NumberValueSetter extends ValueSetter {
		@Override
		protected Set<Ref>[] setCellValue(int row, int col, RefSheet refSheet, Object value) {
			return setCellValue(row, col, refSheet, ((Number)value).doubleValue());
		}
		
		private Set<Ref>[] setCellValue(int rowIndex, int colIndex, RefSheet refSheet, double value) {
			//locate the cell of the refSheet
			final Cell cell = getOrCreateCell(rowIndex, colIndex, refSheet, Cell.CELL_TYPE_NUMERIC);
			return BookHelper.setCellValue(cell, value);
		}
	}
	
	private class BooleanValueSetter extends ValueSetter {
		@Override
		protected Set<Ref>[] setCellValue(int row, int col, RefSheet refSheet, Object value) {
			return setCellValue(row, col, refSheet, ((Boolean)value).booleanValue());
		}
		
		private Set<Ref>[] setCellValue(int rowIndex, int colIndex, RefSheet refSheet, boolean value) {
			//locate the cell of the refSheet
			final Cell cell = getOrCreateCell(rowIndex, colIndex, refSheet, Cell.CELL_TYPE_BOOLEAN);
			return BookHelper.setCellValue(cell, value);
		}
	}
	
	private class ErrorValueSetter extends ValueSetter {
		@Override
		protected Set<Ref>[] setCellValue(int row, int col, RefSheet refSheet, Object value) {
			return setCellValue(row, col, refSheet, ((Byte)value).byteValue());
		}
		
		private Set<Ref>[] setCellValue(int rowIndex, int colIndex, RefSheet refSheet, byte value) {
			//locate the cell of the refSheet
			final Cell cell = getOrCreateCell(rowIndex, colIndex, refSheet, Cell.CELL_TYPE_ERROR);
			return BookHelper.setCellErrorValue(cell, value);
		}
	}
	
	private class DateValueSetter extends ValueSetter {
		@Override
		protected Set<Ref>[] setCellValue(int row, int col, RefSheet refSheet, Object value) {
			return setCellValue(row, col, refSheet, (Date) value);
		}
		
		private Set<Ref>[] setCellValue(int rowIndex, int colIndex, RefSheet refSheet, Date value) {
			//locate the cell of the refSheet
			final Cell cell = getOrCreateCell(rowIndex, colIndex, refSheet, Cell.CELL_TYPE_NUMERIC);
			return BookHelper.setCellValue(cell, value);
		}
	}
	
	private class FormulaValueSetter extends ValueSetter {
		@Override
		protected Set<Ref>[] setCellValue(int row, int col, RefSheet refSheet, Object value) {
			return setCellFormula(row, col, refSheet, (String) value);
		}
		
		private Set<Ref>[] setCellFormula(int rowIndex, int colIndex, RefSheet refSheet, String value) {
			//locate the cell of the refSheet. bug#296: cannot insert =A1 into cell via API
			final Cell cell = getOrCreateCell(rowIndex, colIndex, refSheet, Cell.CELL_TYPE_BLANK);
			final Set<Ref>[] refs = BookHelper.setCellFormula(cell, value);
			if (refs != null && (refs[0] == null || refs[0].isEmpty())) {
				refs[0].add(refSheet.getOrCreateRef(rowIndex, colIndex, rowIndex, colIndex));
			}
			return refs; 
		}
	}
	
	private class StringValueSetter extends ValueSetter {
		@Override
		protected Set<Ref>[] setCellValue(int row, int col, RefSheet refSheet, Object value) {
			return setCellValue(row, col, refSheet, (String) value);
		}
		
		private Set<Ref>[] setCellValue(int rowIndex, int colIndex, RefSheet refSheet, String value) {
			//locate the cell of the refSheet
			final Cell cell = getOrCreateCell(rowIndex, colIndex, refSheet, Cell.CELL_TYPE_STRING);
			return BookHelper.setCellValue(cell, value);
		}
	}
	
	private class RichTextStringValueSetter extends ValueSetter {
		@Override
		protected Set<Ref>[] setCellValue(int row, int col, RefSheet refSheet, Object value) {
			return setCellValue(row, col, refSheet, (RichTextString) value);
		}
		
		private Set<Ref>[] setCellValue(int rowIndex, int colIndex, RefSheet refSheet, RichTextString value) {
			//locate the model book and sheet of the refSheet
			final Cell cell = getOrCreateCell(rowIndex, colIndex, refSheet, Cell.CELL_TYPE_STRING);
			return BookHelper.setCellValue(cell, value);
		}
	}
	
	private Row getRow(int rowIndex, RefSheet refSheet) {
		final Book book = BookHelper.getBook(_sheet, refSheet);
		if (book != null) {
			final Worksheet sheet = book.getWorksheet(refSheet.getSheetName());
			if (sheet != null) {
				return sheet.getRow(rowIndex);
			}
		}
		return null;
	}
	
	private Cell getCell(int rowIndex, int colIndex, RefSheet refSheet) {
		//locate the model book and sheet of the refSheet
		final Book book = BookHelper.getBook(_sheet, refSheet);
		if (book != null) {
			final Worksheet sheet = book.getWorksheet(refSheet.getSheetName());
			if (sheet != null) {
				final Row row = sheet.getRow(rowIndex);
				if (row != null) {
					return row.getCell(colIndex);
				}
			}
		}
		return null;
	}
	
	private Cell getOrCreateCell(int rowIndex, int colIndex, RefSheet refSheet, int cellType) {
		//locate the model book and sheet of the refSheet
		final Book book = BookHelper.getBook(_sheet, refSheet);
		final String sheetname = refSheet.getSheetName();
		final Worksheet sheet = getOrCreateSheet(book, sheetname);
		final Row row = getOrCreateRow(sheet, rowIndex);
		return getOrCreateCell(row, colIndex, cellType);
	}
	
	private Worksheet getOrCreateSheet(Book book, String sheetName) {
		Worksheet sheet = book.getWorksheet(sheetName);
		if (sheet == null) {
			sheet = (Worksheet)book.createSheet(sheetName);
		}
		return sheet;
	}
	
	private Row getOrCreateRow(Worksheet sheet, int rowIndex) {
		Row row = sheet.getRow(rowIndex);
		if (row == null) {
			row = sheet.createRow(rowIndex);
		}
		return row;
	}
	
	private Cell getOrCreateCell(Row row, int colIndex, int type) {
		Cell cell = row.getCell(colIndex);
		if (cell == null) {
			cell = row.createCell(colIndex, type); 
		}
		return cell;
	}

	/*package*/ Ref addRef(Ref ref) {
		final Collection<Ref> refs = getRefs();
		_left = Math.min(ref.getLeftCol(), _left);
		_top = Math.min(ref.getTopRow(), _top);
		_right = Math.max(ref.getRightCol(), _right);
		_bottom = Math.max(ref.getBottomRow(), _bottom);
		refs.add(ref);
		return ref;
	}
	
	@Override
	public void delete(int shift) {
		synchronized(_sheet) {
			if (_refs != null && !_refs.isEmpty()) {
				final Ref ref = _refs.iterator().next();
				final RefSheet refSheet = ref.getOwnerSheet();
				final RefBook refBook = refSheet.getOwnerBook();
				final Worksheet sheet = BookHelper.getSheet(_sheet, refSheet);
				if (!((SheetCtrl)sheet).isEvalAll()) {
					((SheetCtrl)sheet).evalAll();
				}
				switch(shift) {
				default:
				case SHIFT_DEFAULT:
					if (ref.isWholeRow()) {
						final ChangeInfo info = BookHelper.deleteRows(sheet, ref.getTopRow(), ref.getRowCount());
						notifyMergeChange(refBook, info, ref, SSDataEvent.ON_RANGE_DELETE, SSDataEvent.MOVE_V);
					} else if (ref.isWholeColumn()) {
						final ChangeInfo info = BookHelper.deleteColumns(sheet, ref.getLeftCol(), ref.getColumnCount());
						notifyMergeChange(refBook, info, ref, SSDataEvent.ON_RANGE_DELETE, SSDataEvent.MOVE_H);
					}
					break;
				case SHIFT_LEFT:
					if (ref.isWholeRow() || ref.isWholeColumn()) {
						delete(SHIFT_DEFAULT);
					} else {
						final ChangeInfo info = BookHelper.deleteRange(sheet, ref.getTopRow(), ref.getLeftCol(), ref.getBottomRow(), ref.getRightCol(), true);
						notifyMergeChange(refBook, info, ref, SSDataEvent.ON_RANGE_DELETE, SSDataEvent.MOVE_H);
					}
					break;
				case SHIFT_UP:
					if (ref.isWholeRow() || ref.isWholeColumn()) {
						delete(SHIFT_DEFAULT);
					} else {
						final ChangeInfo info = BookHelper.deleteRange(sheet, ref.getTopRow(), ref.getLeftCol(), ref.getBottomRow(), ref.getRightCol(), false);
						notifyMergeChange(refBook, info, ref, SSDataEvent.ON_RANGE_DELETE, SSDataEvent.MOVE_V);
					}
					break;
				}
			}
		}
	}
	
	@Override
	public void insert(int shift, int copyOrigin) {
		synchronized(_sheet) {
			if (_refs != null && !_refs.isEmpty()) {
				final Ref ref = _refs.iterator().next();
				final RefSheet refSheet = ref.getOwnerSheet();
				final RefBook refBook = refSheet.getOwnerBook();
				final Worksheet sheet = BookHelper.getSheet(_sheet, refSheet);
				if (!((SheetCtrl)sheet).isEvalAll()) {
					((SheetCtrl)sheet).evalAll();
				}
				switch(shift) {
				default:
				case SHIFT_DEFAULT:
					if (ref.isWholeRow()) {
						final ChangeInfo info = BookHelper.insertRows(sheet, ref.getTopRow(), ref.getRowCount(), copyOrigin);
						notifyMergeChange(refBook, info, ref, SSDataEvent.ON_RANGE_INSERT, SSDataEvent.MOVE_V);
					} else if (ref.isWholeColumn()) {
						final ChangeInfo info = BookHelper.insertColumns(sheet, ref.getLeftCol(), ref.getColumnCount(), copyOrigin);
						notifyMergeChange(refBook, info, ref, SSDataEvent.ON_RANGE_INSERT, SSDataEvent.MOVE_H);
					}
					break;
				case SHIFT_RIGHT:
					if (ref.isWholeRow() || ref.isWholeColumn()) {
						insert(SHIFT_DEFAULT, copyOrigin);
					} else {
						final ChangeInfo info = BookHelper.insertRange(sheet, ref.getTopRow(), ref.getLeftCol(), ref.getBottomRow(), ref.getRightCol(), true, copyOrigin);
						notifyMergeChange(refBook, info, ref, SSDataEvent.ON_RANGE_INSERT, SSDataEvent.MOVE_H);
					}
					break;
				case SHIFT_DOWN:
					if (ref.isWholeRow() || ref.isWholeColumn()) {
						insert(SHIFT_DEFAULT, copyOrigin);
					} else {
						final ChangeInfo info = BookHelper.insertRange(sheet, ref.getTopRow(), ref.getLeftCol(), ref.getBottomRow(), ref.getRightCol(), false, copyOrigin);
						notifyMergeChange(refBook, info, ref, SSDataEvent.ON_RANGE_INSERT, SSDataEvent.MOVE_V);
					}
					break;
				}
			}
		}
	}
	
	private void notifyMergeChange(RefBook refBook, ChangeInfo info, Ref ref, String event, int orient) {
		if (info == null) {
			return;
		}
		final Set<Ref> last = info.getToEval();
		final Set<Ref> all = info.getAffected();
		if (event != null && ref != null) {
			refBook.publish(new SSDataEvent(event, ref, orient));
		}
		//must delete and add in batch, or merge ranges can interfere to each other
		for(MergeChange change : info.getMergeChanges()) {
			final Ref orgMerge = change.getOrgMerge();
			if (orgMerge != null) {
				refBook.publish(new SSDataEvent(SSDataEvent.ON_MERGE_DELETE, orgMerge, orient));
			}
		}
		for(MergeChange change : info.getMergeChanges()) {
			final Ref merge = change.getMerge();
			if (merge != null) {
				refBook.publish(new SSDataEvent(SSDataEvent.ON_MERGE_ADD, merge, orient));
			}
		}
		BookHelper.reevaluateAndNotify((Book) _sheet.getWorkbook(), last, all);
	}

	@Override
	public Range pasteSpecial(int pasteType, int operation, boolean SkipBlanks, boolean transpose) {
		//TODO
		//clipboardRange.pasteSpecial(this, pasteType, pasteOp, skipBlanks, transpose);
		return null;
	}
	
	//enforce lock sequence
	private Worksheet[] getLockSheets(Range dstRange) {
		final int srcIndex = _sheet.getBook().getSheetIndex(_sheet);
		final Worksheet dstSheet = dstRange.getSheet();
		final int dstIndex = dstSheet.getBook().getSheetIndex(dstSheet);
		final Worksheet[] sheets = new Worksheet[2];
		final Worksheet sheet1 = srcIndex > dstIndex ? _sheet : dstSheet;
		final Worksheet sheet2 = srcIndex > dstIndex ? dstSheet : _sheet;
		return new Worksheet[] {sheet1, sheet2};
	}
	
	@Override
	public Range pasteSpecial(Range dstRange, int pasteType, int pasteOp, boolean skipBlanks, boolean transpose) {
		final Worksheet[] sheets = getLockSheets(dstRange); //enforce lock sequence
		synchronized(sheets[0]) { 
			synchronized(sheets[1]) {
				final Ref ref = paste0(dstRange, pasteType, pasteOp, skipBlanks, transpose);
				return ref == null ? null : new RangeImpl(ref, BookHelper.getSheet(_sheet, ref.getOwnerSheet()));
			}
		}
	}
	
	@SuppressWarnings("unchecked")
	@Override
	public void sort(Range rng1, boolean desc1, Range rng2, int type, boolean desc2, Range rng3, boolean desc3, int header, int orderCustom,
			boolean matchCase, boolean sortByRows, int sortMethod, int dataOption1, int dataOption2, int dataOption3) {
		synchronized(_sheet) {
			final Ref key1 = rng1 != null ? ((RangeImpl)rng1).getRefs().iterator().next() : null;
			final Ref key2 = rng2 != null ? ((RangeImpl)rng2).getRefs().iterator().next() : null;
			final Ref key3 = rng3 != null ? ((RangeImpl)rng3).getRefs().iterator().next() : null;
			if (_refs != null && !_refs.isEmpty()) {
				final Ref ref = _refs.iterator().next();
				final int tRow = ref.getTopRow();
				final int lCol = ref.getLeftCol();
				final int bRow = ref.getBottomRow();
				final int rCol = ref.getRightCol();
				final Worksheet sheet = BookHelper.getSheet(_sheet, ref.getOwnerSheet());
				final RefBook refBook = ref.getOwnerSheet().getOwnerBook();
				ChangeInfo info = BookHelper.sort(sheet, tRow, lCol, bRow, rCol, 
									key1, desc1, key2, type, desc2, key3, desc3, header, 
									orderCustom, matchCase, sortByRows, sortMethod, dataOption1, dataOption2, dataOption3);
				if (info == null) {
					info = new ChangeInfo(new HashSet<Ref>(0), new HashSet<Ref>(), new ArrayList<MergeChange>(0));
				}
				info.getAffected().add(ref);
				notifyMergeChange(refBook, info, ref, SSDataEvent.ON_CONTENTS_CHANGE, SSDataEvent.MOVE_NO);
			}
		}
	}
	@Override
	public Range copy(Range dstRange) {
		final Worksheet[] sheets = getLockSheets(dstRange); //enforce lock sequence
		synchronized(sheets[0]) { //enforce lock sequence
			synchronized(sheets[1]) {
				final Ref ref = paste0(dstRange, Range.PASTE_ALL, Range.PASTEOP_NONE, false, false);
				return ref == null ? null : new RangeImpl(ref, BookHelper.getSheet(_sheet, ref.getOwnerSheet()));
			}
		}
	}
	private Ref paste0(Range dstRange, int pasteType, int pasteOp, boolean skipBlanks, boolean transpose) {
		if (_refs != null && !_refs.isEmpty() && !((RangeImpl)dstRange).getRefs().isEmpty()) {
			//check if follow the copy rule
			//destination range allow only one contiguous reference
			if (((RangeImpl)dstRange).getRefs().size() > 1) {
				throw new UiException("Command cannot be used on multiple selections");
			}
			//source range can handle only same rows/columns multiple references
			Iterator<Ref> it = _refs.iterator();
			Ref ref1 = it.next();
			int srcRowCount = ref1.getRowCount();
			int srcColCount = ref1.getColumnCount();
			final Ref dstRef = ((RangeImpl)dstRange).getRefs().iterator().next();
			final Set<Ref> toEval = new HashSet<Ref>();
			final Set<Ref> affected = new HashSet<Ref>();
			final List<MergeChange> mergeChanges = new ArrayList<MergeChange>();
			final ChangeInfo info = new ChangeInfo(toEval, affected, mergeChanges);
			Ref pasteRef = null;
			if (_refs.size() > 1) { //multiple src references
				final SortedMap<Integer, Ref> srcRefs = new TreeMap<Integer, Ref>();
				boolean sameRow = false;
				boolean sameCol = false;
				final int lCol = ref1.getLeftCol();
				final int tRow = ref1.getTopRow();
				final int rCol = ref1.getRightCol();
				final int bRow = ref1.getBottomRow();
				while (it.hasNext()) {
					final Ref ref = it.next();
					if (lCol == ref.getLeftCol() && rCol == ref.getRightCol()) { //same column
						if (sameRow) { //cannot be both sameRow and sameColumn
							throw new UiException("Command cannot be used on multiple selections");
						}
						if (srcRefs.isEmpty()) {
							srcRefs.put(new Integer(tRow), ref1); //sorted on Row
						}
						srcRefs.put(new Integer(ref.getTopRow()), ref);
						sameCol = true;
						srcRowCount += ref.getRowCount();
					} else if (tRow == ref.getTopRow() && bRow == ref.getBottomRow()) { //same row
						if (sameCol) { //cannot be both sameRow and sameColumn
							throw new UiException("Command cannot be used on multiple selections");
						}
						if (srcRefs.isEmpty()) {
							srcRefs.put(Integer.valueOf(lCol), ref1); //sorted on column
						}
						srcRefs.put(Integer.valueOf(ref.getLeftCol()), ref);
						sameRow = true;
						srcColCount += ref.getColumnCount();
					} else { //not the same column or same row
						throw new UiException("Command cannot be used on multiple selections");
					}
				}
				pasteType = pasteType + Range.PASTE_VALUES; //no formula 
				pasteRef = copyMulti(sameRow, srcRefs, srcColCount, srcRowCount, dstRef, pasteType, pasteOp, skipBlanks, transpose, info);
			} else {
				pasteRef = copy(ref1, srcColCount, srcRowCount, dstRange, dstRef, pasteType, pasteOp, skipBlanks, transpose, info);
			}
			if (pasteRef != null) {
				notifyMergeChange(ref1.getOwnerSheet().getOwnerBook(), info, ref1, SSDataEvent.ON_CONTENTS_CHANGE, SSDataEvent.MOVE_NO);
			}
			return pasteRef;
		}
		return null;
	}

	@Override
	public void borderAround(BorderStyle lineStyle, String color) {
		setBorders(BookHelper.BORDER_OUTLINE, lineStyle, color);
	}
	
	@Override
	public void merge(boolean across) {
		synchronized (_sheet) {
			if (_refs != null && !_refs.isEmpty()) {
				final Ref ref = _refs.iterator().next();
				final RefSheet refSheet = ref.getOwnerSheet();
				final RefBook refBook = refSheet.getOwnerBook();
				final int tRow = ref.getTopRow();
				final int lCol = ref.getLeftCol();
				final int bRow = ref.getBottomRow();
				final int rCol = ref.getRightCol();
				final Worksheet sheet = BookHelper.getSheet(_sheet, refSheet);
				final ChangeInfo info = BookHelper.merge(sheet, tRow, lCol, bRow, rCol, across);
				notifyMergeChange(refBook, info, ref, SSDataEvent.ON_CONTENTS_CHANGE, SSDataEvent.MOVE_NO);
			}
		}
	}
	
	@Override
	public void unMerge() {
		synchronized (_sheet) {
			if (_refs != null && !_refs.isEmpty()) {
				final Ref ref = _refs.iterator().next();
				final RefSheet refSheet = ref.getOwnerSheet();
				final RefBook refBook = refSheet.getOwnerBook();
				final int tRow = ref.getTopRow();
				final int lCol = ref.getLeftCol();
				final int bRow = ref.getBottomRow();
				final int rCol = ref.getRightCol();
				final Worksheet sheet = BookHelper.getSheet(_sheet, refSheet);
				final ChangeInfo info = BookHelper.unMerge(sheet, tRow, lCol, bRow, rCol);
				notifyMergeChange(refBook, info, ref, SSDataEvent.ON_CONTENTS_CHANGE, SSDataEvent.MOVE_NO);
			}
		}
	}
	
	@Override
	public Range getCells(int row, int col) {
		synchronized (_sheet) {
			final Ref ref = getRefs().iterator().next();
			final int col1 = ref.getLeftCol() + col;
			final int row1 = ref.getTopRow() + row;
			return new RangeImpl(row1, col1, _sheet, _sheet);
		}
	}
			
	private Ref copyMulti(boolean sameRow, SortedMap<Integer, Ref> srcRefs, int srcColCount, int srcRowCount, Ref dstRef, int pasteType, int pasteOp, boolean skipBlanks, boolean transpose, ChangeInfo info) {
		final int dstColCount = transpose ? dstRef.getRowCount() : dstRef.getColumnCount();
		final int dstRowCount = transpose ? dstRef.getColumnCount() : dstRef.getRowCount();
		
		if ((dstRowCount % srcRowCount) == 0 && (dstColCount % srcColCount) == 0) {
			return copyRefs(sameRow, srcRefs, srcColCount, srcRowCount, dstColCount/srcColCount, dstRowCount/srcRowCount, dstRef, pasteType, pasteOp, skipBlanks, transpose, info);
		} else if (dstColCount == 1 && (dstRowCount % srcRowCount) == 0) {
			return copyRefs(sameRow, srcRefs, srcColCount, srcRowCount, 1, dstRowCount/srcRowCount, dstRef, pasteType, pasteOp, skipBlanks, transpose, info);
		} else if (dstRowCount == 1 && (dstColCount % srcColCount) == 0) {
			return copyRefs(sameRow, srcRefs, srcColCount, srcRowCount, dstColCount/srcColCount, 1, dstRef, pasteType, pasteOp, skipBlanks, transpose, info);
		} else {
			return copyRefs(sameRow, srcRefs, srcColCount, srcRowCount, 1, 1, dstRef, pasteType, pasteOp, skipBlanks, transpose, info);
		}
	}

	private void rowCopyRefs(Map<Integer, Ref> srcRefs, int srcColCount, int srcRowCount, int colRepeat, int rowRepeat, RefSheet dstSheet, int tRow, int bRow, int lCol, int pasteType, int pasteOp, boolean skipBlanks, boolean transpose, ChangeInfo info) {
		int dsttRow = tRow;
		int dstbRow = bRow;
		for(int rr = rowRepeat; rr > 0; --rr) {
			int lCol0 = lCol;
			for (int cr = colRepeat; cr > 0; --cr) {
				for (Entry<Integer, Ref> srcEntry : srcRefs.entrySet()) {
					final Ref srcRef = srcEntry.getValue();
					int rCol0 = lCol0 + (transpose ? srcRef.getRowCount() : srcRef.getColumnCount()) - 1; 
					final Ref dstRef0 = new AreaRefImpl(dsttRow, lCol0, dstbRow, rCol0, dstSheet);
					copyRef(srcRef, 1, 1, dstRef0, pasteType, pasteOp, skipBlanks, false, info);
					lCol0 = rCol0 + 1;
				}
			}
			dsttRow = dstbRow + 1;
			dstbRow += srcRowCount;
		}
	}
	
	private void colCopyRefs(Map<Integer, Ref> srcRefs, int srcColCount, int srcRowCount, int colRepeat, int rowRepeat, RefSheet dstSheet, int lCol, int rCol, int tRow, int pasteType, int pasteOp, boolean skipBlanks, boolean transpose, ChangeInfo info) {
		int dstlCol = lCol;
		int dstrCol = rCol;
		for (int cr = colRepeat; cr > 0; --cr) {
			int tRow0 = tRow;
			for(int rr = rowRepeat; rr > 0; --rr) {
				for (Entry<Integer, Ref> srcEntry : srcRefs.entrySet()) {
					final Ref srcRef = srcEntry.getValue();
					int bRow0 = tRow0 + (transpose ? srcRef.getColumnCount() : srcRef.getRowCount()) - 1; 
					final Ref dstRef0 = new AreaRefImpl(tRow0, dstlCol, bRow0, dstrCol, dstSheet);
					copyRef(srcRef, 1, 1, dstRef0, pasteType, pasteOp, skipBlanks, false, info);
					tRow0 = bRow0 + 1;
				}
			}
			dstlCol += dstrCol + 1;
			dstrCol += srcColCount;
		}
	}
	
	//Any cell is protecetd and locked
	@Override
	public boolean isAnyCellProtected() {
		synchronized (_sheet) {
			final Ref ref = _refs != null && !_refs.isEmpty() ? _refs.iterator().next() : null;
			if (ref != null) {
				final RefSheet refSheet = ref.getOwnerSheet();
				final int left = ref.getLeftCol();
				final int right = ref.getRightCol();
				final int top = ref.getTopRow();
				final int bottom = ref.getBottomRow();
				
				//ZSS-22: Shall not allow Copy and Paste operation in a protected spreadsheet
				final Worksheet sheet = BookHelper.getSheet(_sheet, refSheet);
				if (sheet.getProtect()) {
					for (int r = top; r <= bottom; ++r) {
						final Row row = sheet.getRow(r);
						if (row != null) {
							for (int c = left; c <= right; ++c) {
								final Cell cell = row.getCell(c);
								if (cell != null) {
									final CellStyle cs = cell.getCellStyle();
									if (cs != null && cs.getLocked()) {
										//as long as one is protected and locked, return true
										return true;
									}
								}
							}
						}
					}
				}
				return false;
			}
			return true;
		}
	}
	
	private Ref getPasteRef(int srcRowCount, int srcColCount, int rowRepeat, int colRepeat, Ref dstRef, boolean transpose) {
		final RefSheet dstRefSheet = dstRef.getOwnerSheet();
		final int dstT = dstRef.getTopRow();
		final int dstL = dstRef.getLeftCol();
		final int dstRowCount = transpose ? srcColCount * colRepeat : srcRowCount * rowRepeat;
		final int dstColCount = transpose ? srcRowCount * rowRepeat : srcColCount * colRepeat; 
		final int dstB = dstT + dstRowCount - 1;
		final int dstR = dstL + dstColCount - 1;
		
		//ZSS-22: Shall not allow Copy and Paste operation in a protected spreadsheet
		final Worksheet dstSheet = BookHelper.getSheet(_sheet, dstRefSheet);
		final Ref pasteRef = new AreaRefImpl(dstT, dstL, dstB, dstR, dstRefSheet);
		return new RangeImpl(pasteRef, dstSheet).isAnyCellProtected() ? null : pasteRef;
	}
	private Ref copyRefs(boolean sameRow, SortedMap<Integer, Ref> srcRefs, int srcColCount, int srcRowCount, int colRepeat, int rowRepeat, Ref dstRef, int pasteType, int pasteOp, boolean skipBlanks, boolean transpose, ChangeInfo info) {
		final Ref pasteRef = getPasteRef(srcRowCount, srcColCount, rowRepeat, colRepeat, dstRef, transpose);
		if (pasteRef == null) {
			return null;
		}
		if (pasteType == Range.PASTE_COLUMN_WIDTHS) {
			final Integer lastKey = srcRefs.lastKey();
			final Ref srcRef = srcRefs.get(lastKey);
			final int srclCol = srcRef.getLeftCol();
			final Worksheet srcSheet = BookHelper.getSheet(_sheet, srcRef.getOwnerSheet());
			final int widthRepeatCount = srcRef.getColumnCount();
			copyColumnWidths(srcSheet, widthRepeatCount, srclCol, pasteRef);
			return pasteRef;
		}
		
		final RefSheet dstSheet = dstRef.getOwnerSheet();
		final int tRow = dstRef.getTopRow();
		final int lCol = dstRef.getLeftCol();
		if (!transpose) {
			final int bRow = tRow + srcRowCount - 1;
			final int rCol = lCol + srcColCount - 1;
			if (sameRow) {
				rowCopyRefs(srcRefs, srcColCount, srcRowCount, colRepeat, rowRepeat, dstSheet, tRow, bRow, lCol, pasteType, pasteOp, skipBlanks, transpose, info);
			} else { //sameCol
				colCopyRefs(srcRefs, srcColCount, srcRowCount, colRepeat, rowRepeat, dstSheet, lCol, rCol, tRow, pasteType, pasteOp, skipBlanks, transpose, info);
			}
		} else { //row -> column, column -> row
			final int bRow = tRow + srcColCount - 1;
			final int rCol = lCol + srcRowCount - 1;
			if (sameRow) {
				colCopyRefs(srcRefs, srcRowCount, srcColCount, colRepeat, rowRepeat, dstSheet, lCol, rCol, tRow, pasteType, pasteOp, skipBlanks, transpose, info);
			} else { //sameCol
				rowCopyRefs(srcRefs, srcRowCount, srcColCount, colRepeat, rowRepeat, dstSheet, tRow, bRow, lCol, pasteType, pasteOp, skipBlanks, transpose, info);
			}
		}
		return pasteRef;
	}
	
	private Ref copy(Ref srcRef, int srcColCount, int srcRowCount, Range dstRange, Ref dstRef, int pasteType, int pasteOp, boolean skipBlanks, boolean transpose, ChangeInfo info) {
		if (pasteType == Range.PASTE_COLUMN_WIDTHS) { //ignore transpose in such case when only one srcRef
			transpose = false; //ignore transpose
		}
		final int dstColCount = transpose ? dstRef.getRowCount() : dstRef.getColumnCount();
		final int dstRowCount = transpose ? dstRef.getColumnCount() : dstRef.getRowCount();
		
		if ((dstRowCount % srcRowCount) == 0 && (dstColCount % srcColCount) == 0) {
			return copyRef(srcRef, dstColCount/srcColCount, dstRowCount/srcRowCount, dstRef, pasteType, pasteOp, skipBlanks, transpose, info);
		} else if (dstColCount == 1 && (dstRowCount % srcRowCount) == 0) {
			return copyRef(srcRef, 1, dstRowCount/srcRowCount, dstRef, pasteType, pasteOp, skipBlanks, transpose, info);
		} else if (dstRowCount == 1 && (dstColCount % srcColCount) == 0) {
			return copyRef(srcRef, dstColCount/srcColCount, 1, dstRef, pasteType, pasteOp, skipBlanks, transpose, info);
		} else {
			return copyRef(srcRef, 1, 1, dstRef, pasteType, pasteOp, skipBlanks, transpose, info);
		}
	}

	private void copyColumnWidths(Worksheet srcSheet, int widthRepeatCount, int srclCol, Ref dstRef) {
		final int dstlCol = dstRef.getLeftCol();
		final int dstColCount = dstRef.getColumnCount();
		final RefSheet dstRefSheet = dstRef.getOwnerSheet();
		final Worksheet dstSheet = BookHelper.getSheet(_sheet, dstRefSheet);
		for (int count = 0; count < dstColCount; ++count) {
			final int dstCol = dstlCol + count;
			final int srcCol = srclCol + count % widthRepeatCount; 
			final int char256 = srcSheet.getColumnWidth(srcCol);
			BookHelper.setColumnWidth(dstSheet, dstCol, dstCol, char256);
		}
		final Book book = (Book) dstSheet.getWorkbook();
		//bug# ZSS-52: Past special, copy column width's behavior doesn't correct
		final int maxrow = book.getSpreadsheetVersion().getLastRowIndex();
		final Set<Ref> affected = new HashSet<Ref>();
		affected.add(dstRef.isWholeColumn() ? dstRef : new AreaRefImpl(0, dstlCol, maxrow, dstRef.getRightCol(), dstRefSheet));
		BookHelper.notifySizeChanges(book, affected);
	}
	
	private Ref copyRef(Ref srcRef, int colRepeat, int rowRepeat, Ref dstRef, int pasteType, int pasteOp, boolean skipBlanks, boolean transpose, ChangeInfo info) {
		final int srcRowCount = srcRef.getRowCount();
		final int srcColCount = srcRef.getColumnCount();
		final Ref pasteRef = getPasteRef(srcRowCount, srcColCount, rowRepeat, colRepeat, dstRef, transpose);
		if (pasteRef == null) {
			return null;
		}
		if (pasteType == Range.PASTE_COLUMN_WIDTHS) {
			final int srclCol = srcRef.getLeftCol();
			final Worksheet srcSheet = BookHelper.getSheet(_sheet, srcRef.getOwnerSheet());
			copyColumnWidths(srcSheet, srcColCount, srclCol, pasteRef);
			return pasteRef;
		}
		final int tRow = srcRef.getTopRow();
		final int lCol = srcRef.getLeftCol();
		final int bRow = srcRef.getBottomRow();
		final int rCol = srcRef.getRightCol();
		final RefSheet dstRefsheet = dstRef.getOwnerSheet();
		final Worksheet srcSheet = BookHelper.getSheet(_sheet, srcRef.getOwnerSheet());
		final Worksheet dstSheet = BookHelper.getSheet(_sheet, dstRefsheet);
		final Set<Ref> toEval = info.getToEval();
		final Set<Ref> affected = info.getAffected();
		final List<MergeChange> mergeChanges = info.getMergeChanges();
		if (!transpose) {
			int dstRow = dstRef.getTopRow();
			for(int rr = rowRepeat; rr > 0; --rr) {
				for(int srcRow = tRow; srcRow <= bRow; ++srcRow, ++dstRow) {
					int dstCol= dstRef.getLeftCol();
					for (int cr = colRepeat; cr > 0; --cr) {
						for (int srcCol = lCol; srcCol <= rCol; ++srcCol, ++dstCol) {
							final Cell cell = BookHelper.getCell(srcSheet, srcRow, srcCol);
							if (cell != null) {
								if (!skipBlanks || cell.getCellType() != Cell.CELL_TYPE_BLANK) {
									final ChangeInfo changeInfo0 = BookHelper.copyCell(cell, dstSheet, dstRow, dstCol, pasteType, pasteOp, transpose);
									BookHelper.assignChangeInfo(toEval, affected, mergeChanges, changeInfo0);
								}
							} else if (!skipBlanks) {
								final Set<Ref>[] refs = BookHelper.removeCell(dstSheet, dstRow, dstCol);
								BookHelper.assignRefs(toEval, affected, refs);
							}
						}
					}
				}
			}
		} else { //row -> column, column -> row
			int dstCol = dstRef.getLeftCol(); 
			for(int rr = rowRepeat; rr > 0; --rr) {
				for(int srcRow = tRow; srcRow <= bRow; ++srcRow, ++dstCol) {
					int dstRow = dstRef.getTopRow();
					for (int cr = colRepeat; cr > 0; --cr) {
						for (int srcCol = lCol; srcCol <= rCol; ++srcCol, ++dstRow) {
							final Cell cell = BookHelper.getCell(srcSheet, srcRow, srcCol);
							if (cell != null) {
								if (!skipBlanks || cell.getCellType() != Cell.CELL_TYPE_BLANK) {
									final ChangeInfo changeInfo0 = BookHelper.copyCell(cell, dstSheet, dstRow, dstCol, pasteType, pasteOp, transpose);
									BookHelper.assignChangeInfo(toEval, affected, mergeChanges, changeInfo0);
								}
							} else if (!skipBlanks) {
								final Set<Ref>[] refs = BookHelper.removeCell(dstSheet, dstRow, dstCol);
								BookHelper.assignRefs(toEval, affected, refs);
							}
						}
					}
				}
			}
		}
		affected.add(pasteRef);
		return pasteRef;
	}
	
	@Override
	public void setBorders(short borderIndex, BorderStyle lineStyle, String color) {
		synchronized (_sheet) {
			if (_refs != null && !_refs.isEmpty()) {
				final Ref ref = _refs.iterator().next();
				final RefSheet refSheet = ref.getOwnerSheet();
				final int tRow = ref.getTopRow();
				final int lCol = ref.getLeftCol();
				final int bRow = ref.getBottomRow();
				final int rCol = ref.getRightCol();
				final Worksheet sheet = BookHelper.getSheet(_sheet, refSheet);
				final Set<Ref> all = BookHelper.setBorders(sheet, tRow, lCol, bRow, rCol, borderIndex, lineStyle, color);
				if (all != null) {
					final Book book = (Book) _sheet.getWorkbook();
					BookHelper.notifyCellChanges(book, all);
				}
			}
		}
	}
	
	@Override
	public void setColumnWidth(int char256) {
		synchronized (_sheet) {
			if (_refs != null && !_refs.isEmpty()) {
				final Ref ref = _refs.iterator().next();
				final RefSheet refSheet = ref.getOwnerSheet();
				final int lCol = ref.getLeftCol();
				final int rCol = ref.getRightCol();
				final Worksheet sheet = BookHelper.getSheet(_sheet, refSheet);
				final Set<Ref> all = BookHelper.setColumnWidth(sheet, lCol, rCol, char256);
				if (all != null) {
					final Book book = (Book) _sheet.getWorkbook();
					BookHelper.notifySizeChanges(book, all);
				}
			}
		}
	}
	
	@Override
	public void setRowHeight(int points) {
		setRowHeight(points, true);
	}
	
	@Override
	public void setRowHeight(int points, boolean customHeight) {
		synchronized (_sheet) {
			if (_refs != null && !_refs.isEmpty()) {
				final Ref ref = _refs.iterator().next();
				final RefSheet refSheet = ref.getOwnerSheet();
				final int tRow = ref.getTopRow();
				final int bRow = ref.getBottomRow();
				final Worksheet sheet = BookHelper.getSheet(_sheet, refSheet);
				final Set<Ref> all = BookHelper.setRowHeight(sheet, tRow, bRow, (short) (points * 20), customHeight); //in twips, set customHeight
				if (all != null) {
					final Book book = (Book) _sheet.getWorkbook();
					BookHelper.notifySizeChanges(book, all);
				}
			}
		}
	}
	
	@Override
	public void move(int nRow, int nCol) {
		synchronized (_sheet) {
			if (_refs != null && !_refs.isEmpty()) {
				final Ref ref = _refs.iterator().next();
				final RefSheet refSheet = ref.getOwnerSheet();
				final RefBook refBook = refSheet.getOwnerBook();
				final int tRow = ref.getTopRow();
				final int lCol = ref.getLeftCol();
				final int bRow = ref.getBottomRow();
				final int rCol = ref.getRightCol();
				final Worksheet sheet = BookHelper.getSheet(_sheet, refSheet);
				final ChangeInfo info = BookHelper.moveRange(sheet, tRow, lCol, bRow, rCol, nRow, nCol);
				notifyMergeChange(refBook, info, ref, SSDataEvent.ON_CONTENTS_CHANGE, SSDataEvent.MOVE_NO);
			}
		}
	}
	
	@Override
	public void setStyle(CellStyle style) {
		synchronized (_sheet) {
			if (_refs != null && !_refs.isEmpty()) {
				final Set<Ref> all = new HashSet<Ref>();
				for (Ref ref : _refs) {
					final RefSheet refSheet = ref.getOwnerSheet();
					final int tRow = ref.getTopRow();
					final int lCol = ref.getLeftCol();
					final int bRow = ref.getBottomRow();
					final int rCol = ref.getRightCol();
					final Worksheet sheet = BookHelper.getSheet(_sheet, refSheet);
					final Set<Ref> refs = BookHelper.setCellStyle(sheet, tRow, lCol, bRow, rCol, style);
					all.addAll(refs);
				}
				if (!all.isEmpty()) {
					final Book book = (Book) _sheet.getWorkbook();
					BookHelper.notifyCellChanges(book, all);
				}
			}
		}
	}
	
	@Override
	public void autoFill(Range dstRange, int fillType) {
		synchronized (_sheet) {
			if (_refs != null && !_refs.isEmpty() && !((RangeImpl)dstRange).getRefs().isEmpty()) {
				//destination range allow only one contiguous reference
				if (((RangeImpl)dstRange).getRefs().size() > 1) {
					throw new UiException("Command cannot be used on multiple selections");
				}
				final Ref srcRef = _refs.iterator().next();
				final Ref dstRef = ((RangeImpl)dstRange).getRefs().iterator().next();
				fillRef(srcRef, dstRef, fillType);
			}
		}
	}
	
	@Override
	public void clearContents() {
		synchronized (_sheet) {
			if (_refs != null && !_refs.isEmpty()) {
				final Ref ref = _refs.iterator().next();
				final RefSheet refSheet = ref.getOwnerSheet();
				final int tRow = ref.getTopRow();
				final int lCol = ref.getLeftCol();
				final int bRow = ref.getBottomRow();
				final int rCol = ref.getRightCol();
				final Worksheet sheet = BookHelper.getSheet(_sheet, refSheet);
				clearContents(sheet, tRow, lCol, bRow, rCol);
			}
		}
	}
	
	private void clearContents(Worksheet sheet, int tRow, int lCol, int bRow, int rCol) {
		final Set<Ref> last = new HashSet<Ref>();
		final Set<Ref> all = new HashSet<Ref>();
		for(int r = tRow; r <= bRow; ++r) {
			for(int c = lCol; c <= rCol; ++c) {
				final Set<Ref>[] refs = BookHelper.clearCell(sheet, r, c);
				if (refs != null) {
					last.addAll(refs[0]);
					all.addAll(refs[1]);
				}
			}
		}
		final Book book = (Book) sheet.getWorkbook();
		all.add(new AreaRefImpl(tRow, lCol, bRow, rCol, BookHelper.getRefSheet(book, sheet)));
		BookHelper.reevaluateAndNotify(book, last, all);
	}
	
	private void fillRef(Ref srcRef, Ref dstRef, int fillType) {
		final ChangeInfo info = BookHelper.fill(_sheet, srcRef, dstRef, fillType);
		notifyMergeChange(dstRef.getOwnerSheet().getOwnerBook(), info, dstRef, SSDataEvent.ON_CONTENTS_CHANGE, SSDataEvent.MOVE_NO);
	}

	@Override
	public void fillDown() {
		synchronized (_sheet) {
			if (_refs != null && !_refs.isEmpty()) {
				final Ref dstRef = _refs.iterator().next();
				final Ref srcRef = new AreaRefImpl(dstRef.getTopRow(), dstRef.getLeftCol(), dstRef.getTopRow(), dstRef.getRightCol(), dstRef.getOwnerSheet());
				fillRef(srcRef, dstRef, Range.FILL_COPY);
			}
		}
	}
	
	@Override
	public void fillLeft() {
		synchronized (_sheet) {
			if (_refs != null && !_refs.isEmpty()) {
				final Ref dstRef = _refs.iterator().next();
				final Ref srcRef = new AreaRefImpl(dstRef.getTopRow(), dstRef.getRightCol(), dstRef.getBottomRow(), dstRef.getRightCol(), dstRef.getOwnerSheet());
				fillRef(srcRef, dstRef, Range.FILL_COPY);
			}
		}
	}

	@Override
	public void fillRight() {
		synchronized (_sheet) {
			if (_refs != null && !_refs.isEmpty()) {
				final Ref dstRef = _refs.iterator().next();
				final Ref srcRef = new AreaRefImpl(dstRef.getTopRow(), dstRef.getLeftCol(), dstRef.getBottomRow(), dstRef.getLeftCol(), dstRef.getOwnerSheet());
				fillRef(srcRef, dstRef, Range.FILL_COPY);
			}
		}
	}

	@Override
	public void fillUp() {
		synchronized (_sheet) {
			if (_refs != null && !_refs.isEmpty()) {
				final Ref dstRef = _refs.iterator().next();
				final Ref srcRef = new AreaRefImpl(dstRef.getBottomRow(), dstRef.getLeftCol(), dstRef.getBottomRow(), dstRef.getRightCol(), dstRef.getOwnerSheet());
				fillRef(srcRef, dstRef, Range.FILL_COPY);
			}
		}
	}
	
	@Override
	public void setHidden(boolean hidden) {
		synchronized (_sheet) {
			if (_refs != null && !_refs.isEmpty()) {
				final Set<Ref> all = new HashSet<Ref>();
				for (Ref ref : _refs) {
					if (ref.isWholeRow()) {
						final RefSheet refSheet = ref.getOwnerSheet();
						final Worksheet sheet = BookHelper.getSheet(_sheet, refSheet);
						final int tRow = ref.getTopRow();
						final int bRow = ref.getBottomRow();
						final Set<Ref> refs = BookHelper.setRowsHidden(sheet, tRow, bRow, hidden);
						all.addAll(refs);
					} else if (ref.isWholeColumn()) {
						final RefSheet refSheet = ref.getOwnerSheet();
						final Worksheet sheet = BookHelper.getSheet(_sheet, refSheet);
						final int lCol = ref.getLeftCol();
						final int rCol = ref.getRightCol();
						final Set<Ref> refs = BookHelper.setColumnsHidden(sheet, lCol, rCol, hidden);
						all.addAll(refs);
					}
				}
				if (!all.isEmpty()) {
					final Book book = (Book) _sheet.getWorkbook();
					BookHelper.notifySizeChanges(book, all);
				}
			}
		}
	}
	
	@Override
	public void setDisplayGridlines(boolean show) {
		synchronized (_sheet) {
			final Ref ref = getRefs().iterator().next();
			final Worksheet sheet = BookHelper.getSheet(_sheet, ref.getOwnerSheet());
			final Set<Ref> all = new HashSet<Ref>(); 
			final boolean old = sheet.isDisplayGridlines();
			if (old != show) {
				sheet.setDisplayGridlines(show);
				//sheet is important, row, column is not in this event
				all.add(ref); 
			}
			if (!all.isEmpty()) {
				final Book book = (Book) sheet.getWorkbook();
				BookHelper.notifyGridlines(book, all, show);
			}
		}
	}
	
	@Override
	public void protectSheet(String password) {
		synchronized (_sheet) {
			final Ref ref = getRefs().iterator().next();
			final Worksheet sheet = BookHelper.getSheet(_sheet, ref.getOwnerSheet());
			final Set<Ref> all = new HashSet<Ref>(); 
			final boolean oldProtected = sheet.getProtect();
			if (oldProtected && password == null) {
				sheet.protectSheet(null);
				all.add(ref);
			} else if (oldProtected == false && password != null) {
				sheet.protectSheet(password);
				all.add(ref);
			}
			if (!all.isEmpty()) {
				final Book book = (Book) sheet.getWorkbook();
				BookHelper.notifyProtectSheet(book, all, password);
			}
		}
	}

	@Override
	public Areas getAreas() {
		final AreasImpl areas = new AreasImpl();
		if (getRefs().size() == 1) {
			areas.addArea(this);
		} else {
			for(Ref ref : getRefs()) {
				final Range rng = refToRange(ref);
				areas.addArea(rng);
			}
		}
		return areas;
	}
	
	private Range refToRange(Ref ref) {
		return new RangeImpl(ref, _sheet);
	}

	@Override
	public long getCount() {
		final Ref ref = getRefs().iterator().next();
		final int col1 = ref.getLeftCol();
		final int col2 = ref.getRightCol();
		final int row1 = ref.getTopRow();
		final int row2 = ref.getBottomRow();
		if (ref.isWholeColumn()) {
			return (long) (col2 - col1 + 1);
		} else if (ref.isWholeRow()) {
			return (long) (row2 - row1 + 1);
		} else if (ref.isWholeSheet()) {
			return 1L;
		}
		final int ccount = col2 - col1 + 1;
		final long rcount = (long) (col2 - col1 + 1);
		return rcount * ccount;
	}
	
	@Override
	public Range getColumns() {
		final Ref ref = getRefs().iterator().next();
		final Worksheet sheet = BookHelper.getSheet(_sheet, ref.getOwnerSheet());
		final int col1 = ref.getLeftCol();
		final int col2 = ref.getRightCol();
		final Book book = (Book) sheet.getWorkbook();
		return new RangeImpl(0, col1, book.getSpreadsheetVersion().getLastRowIndex(), col2, sheet, sheet);
	}
	
	@Override
	public Range getRows() {
		final Ref ref = getRefs().iterator().next();
		final Worksheet sheet = BookHelper.getSheet(_sheet, ref.getOwnerSheet());
		final int row1 = ref.getTopRow();
		final int row2 = ref.getBottomRow();
		final Book book = (Book) sheet.getWorkbook();
		return new RangeImpl(row1, 0, row2, book.getSpreadsheetVersion().getLastColumnIndex(), sheet, sheet);
	}
	
	@Override
	public Range getDependents() {
		synchronized (_sheet) {
			final Ref ref = getRefs().iterator().next();
			final Worksheet sheet = BookHelper.getSheet(_sheet, ref.getOwnerSheet());
			final int row = ref.getTopRow();
			final int col = ref.getLeftCol();
			final RefSheet refSheet = ref.getOwnerSheet();
			Set<Ref> refs = ((RefSheetImpl)refSheet).getAllDependents(row, col);
			return refs != null && !refs.isEmpty() ?
					new RangeImpl(refs, sheet) : Ranges.EMPTY_RANGE;
		}
	}
	
	@Override
	public Range getDirectDependents() {
		synchronized (_sheet) {
			final Ref ref = getRefs().iterator().next();
			final Worksheet sheet = BookHelper.getSheet(_sheet, ref.getOwnerSheet());
			final int row = ref.getTopRow();
			final int col = ref.getLeftCol();
			final RefSheet refSheet = ref.getOwnerSheet();
			Set<Ref> refs = ((RefSheetImpl)refSheet).getDirectDependents(row, col);
			return refs != null && !refs.isEmpty() ?
					new RangeImpl(refs, sheet) : Ranges.EMPTY_RANGE;
		}
	}
	
	@Override
	public Range getPrecedents() {
		synchronized (_sheet) {
			final Ref ref = getRefs().iterator().next();
			final Worksheet sheet = BookHelper.getSheet(_sheet, ref.getOwnerSheet());
			final int row = ref.getTopRow();
			final int col = ref.getLeftCol();
			final RefSheet refSheet = ref.getOwnerSheet();
			Set<Ref> refs = refSheet.getAllPrecedents(row, col);
			return refs != null && !refs.isEmpty() ?
					new RangeImpl(refs, sheet) : Ranges.EMPTY_RANGE;
		}
		
	}
	
	@Override
	public Range getDirectPrecedents() {
		synchronized (_sheet) {
			final Ref ref = getRefs().iterator().next();
			final Worksheet sheet = BookHelper.getSheet(_sheet, ref.getOwnerSheet());
			final int row = ref.getTopRow();
			final int col = ref.getLeftCol();
			final RefSheet refSheet = ref.getOwnerSheet();
			Set<Ref> refs = refSheet.getDirectPrecedents(row, col);
			return refs != null && !refs.isEmpty() ?
					new RangeImpl(refs, sheet) : Ranges.EMPTY_RANGE;
		}
	}
	
	@Override
	public int getRow() {
		final Ref ref = getRefs().iterator().next();
		return ref.getTopRow();
	}
	
	@Override
	public int getColumn() {
		final Ref ref = getRefs().iterator().next();
		return ref.getLeftCol();
	}

	@Override
	public int getLastColumn() {
		final Ref ref = getRefs().iterator().next();
		return ref.getRightCol();
	}

	@Override
	public int getLastRow() {
		final Ref ref = getRefs().iterator().next();
		return ref.getBottomRow();
	}

	@Override
	public Object getValue() {
		synchronized (_sheet) {
			Ref ref = _refs != null && !_refs.isEmpty() ? _refs.iterator().next() : null;
			if (ref != null) {
				final int tRow = ref.getTopRow();
				final int lCol = ref.getLeftCol();
				final RefSheet refSheet = ref.getOwnerSheet();
				final Cell cell = getCell(tRow, lCol, refSheet);
				if (cell != null) {
					return getValue0(cell);
				}
			}
			return null;
		}
	}

	private Object getValue0(Cell cell) {
		int cellType = cell.getCellType();
		if (cellType == Cell.CELL_TYPE_FORMULA) {
			final Book book = (Book)cell.getSheet().getWorkbook();
			final CellValue cv = BookHelper.evaluate(book, cell);
			return BookHelper.getValueByCellValue(cv);
		} else {
			final Object obj = BookHelper.getCellValue(cell);
			return obj instanceof RichTextString ?
					((RichTextString)obj).getString() : obj;
		}
	}

	@Override
	public void setValue(Object value) {
		//ZSS-78: Implementation of RangeImpl#setValue() is not correct
		synchronized (_sheet) {
			Ref ref = _refs != null && !_refs.isEmpty() ? _refs.iterator().next() : null;
			if (ref != null) {
				Set<Ref>[] refs = null;
				if (value instanceof FormulaError) {
					refs = setValue(Byte.valueOf(((FormulaError)value).getCode()));
				} else if (value instanceof Byte) {
					try {
						FormulaError.forInt(((Byte)value).byteValue());
						refs = setValue((Byte) value);
					} catch (IllegalArgumentException ex) {
						refs = setValue((Number) value);
					}
				} else if (value instanceof Number) {
					refs = setValue((Number)value);
				} else if (value instanceof String) {
					refs = setValue((String) value);
				} else if (value instanceof Boolean) {
					refs = setValue((Boolean) value);
				} else if (value instanceof Date) {
					refs = setValue((Date) value);
				} else if (value == null) {
					refs = setValue((String) null);
				} else {
					 throw new IllegalArgumentException("Unknown value type; must be a FormulaError, a leagal errorCode(a Byte), a Number, a String, a Boolean, a Date, or null: " + value);
				}
				if (refs != null) {
					reevaluateAndNotify(refs);
				}
			}
		}
	}

	@Override
	public Range getOffset(int rowOffset, int colOffset) {
		if (rowOffset == 0 && colOffset == 0) { //no offset, return this
			return this;
		}
		if (_refs != null && !_refs.isEmpty()) {
			final SpreadsheetVersion ver = ((Book)_sheet.getWorkbook()).getSpreadsheetVersion();
			final int maxCol = ver.getLastColumnIndex();
			final int maxRow = ver.getLastRowIndex();
			final Set<Ref> nrefs = new LinkedHashSet<Ref>(_refs.size());
			final Map<RefAddr, Ref> refMap = new HashMap<RefAddr, Ref>(_refs.size()); //index of Ref left/top/right/bottom 

			for(Ref ref : _refs) {
				final int left = ref.getLeftCol() + colOffset;
				final int top = ref.getTopRow() + rowOffset;
				final int right = ref.getRightCol() + colOffset;
				final int bottom = ref.getBottomRow() + rowOffset;
				
				final RefSheet refSheet = ref.getOwnerSheet();
				final int nleft = colOffset < 0 ? Math.max(0, left) : left;  
				final int ntop = rowOffset < 0 ? Math.max(0, top) : top;
				final int nright = colOffset > 0 ? Math.min(maxCol, right) : right;
				final int nbottom = rowOffset > 0 ? Math.min(maxRow, bottom) : bottom;
				
				if (nleft > nright || ntop > nbottom) { //offset out of range
					continue;
				}
				final RefAddr refAddr = new RefAddr(ntop, nleft, nbottom, nright);
				if (refMap.containsKey(refAddr)) { //same area there, next
					continue;
				}
				final Ref newRef = (nleft == nright && ntop == nbottom) ? 
					new CellRefImpl(ntop, nleft, refSheet) :
					new AreaRefImpl(ntop, nleft, nbottom, nright, refSheet);
				nrefs.add(newRef);
				refMap.put(refAddr, newRef);
			}
			if (nrefs.isEmpty()) {
				return Ranges.EMPTY_RANGE;
			} else{
				return new RangeImpl(nrefs, _sheet);
			}
		}
		return Ranges.EMPTY_RANGE;
	}

	//returns the largest square range of this sheet that contains non-blank cells
	private CellRangeAddress getLargestRange(Worksheet sheet) {
		int t = sheet.getFirstRowNum();
		int b = sheet.getLastRowNum();
		//top row
		int minr = -1;
		for(int r = t; r <= b && minr < 0; ++r) {
			final Row rowobj = sheet.getRow(r);
			if (rowobj != null) {
				int ll = rowobj.getFirstCellNum();
				if (ll < 0) { //empty row
					continue;
				}
				int rr = rowobj.getLastCellNum() - 1;
				for(int c = ll; c <= rr; ++c) {
					final Cell cell = rowobj.getCell(c);
					if (!BookHelper.isBlankCell(cell)) { //first no blank row
						minr = r;
						break;
					}
				}
			}
		}
		//bottom row
		int maxr = -1;
		for(int r = b; r >= minr && maxr < 0; --r) {
			final Row rowobj = sheet.getRow(r);
			if (rowobj != null) {
				int ll = rowobj.getFirstCellNum();
				if (ll < 0) { //empty row
					continue;
				}
				int rr = rowobj.getLastCellNum() - 1;
				for(int c = ll; c <= rr; ++c) {
					final Cell cell = rowobj.getCell(c);
					if (!BookHelper.isBlankCell(cell)) { //first no blank row
						maxr = r;
						break;
					}
				}
			}
		}
		//left col
		int minc = Integer.MAX_VALUE;
		for(int r = minr; r <= maxr; ++r) {
			final Row rowobj = sheet.getRow(r);
			if (rowobj != null) {
				int ll = rowobj.getFirstCellNum();
				if (ll < 0) { //empty row
					continue;
				}
				int rr = rowobj.getLastCellNum() - 1;
				for(int c = ll; c < minc && c <= rr; ++c) {
					final Cell cell = rowobj.getCell(c);
					if (!BookHelper.isBlankCell(cell)) { //first no blank row
						minc = c;
						break;
					}
				}
			}
		}
		//right col
		int maxc = -1;
		for(int r = minr; r <= maxr; ++r) {
			final Row rowobj = sheet.getRow(r);
			if (rowobj != null) {
				int ll = rowobj.getFirstCellNum();
				if (ll < 0) { //empty row
					continue;
				}
				int rr = rowobj.getLastCellNum() - 1;
				for(int c = rr; c > maxc && c >= ll; --c) {
					final Cell cell = rowobj.getCell(c);
					if (!BookHelper.isBlankCell(cell)) { //first no blank row
						maxc = c;
						break;
					}
				}
			}
		}
		
		if (minr < 0 || maxc < 0) { //all blanks!
			return null;
		}
		return new CellRangeAddress(minr, maxr, minc, maxc);
	}
	
	@Override
	public Range getCurrentRegion() {
		synchronized (_sheet) {
			final Ref ref = getRefs().iterator().next();
			final int row = ref.getTopRow();
			final int col = ref.getLeftCol();
			CellRangeAddress cra = getCurrentRegion(_sheet, row, col);
			return cra == null ? 
					new RangeImpl(row, col, _sheet, _sheet) :
					new RangeImpl(cra.getFirstRow(), cra.getFirstColumn(), cra.getLastRow(), cra.getLastColumn(), _sheet, _sheet);
		}
	}
	
	private int[] getCellMinMax(Worksheet sheet, int row, int col) {
		final CellRangeAddress rng = BookHelper.getMergeRegion(sheet, row, col);
		final int t = rng.getFirstRow();
		final int l = rng.getFirstColumn();
		final int b = rng.getLastRow();
		final int r = rng.getLastColumn();
		final Cell cell = BookHelper.getCell(sheet, t, l);
		return (!BookHelper.isBlankCell(cell)) ?
			new int[] {l, t, r, b} : null;
	}
	
	//[0]: left, [1]: top, [2]: right, [3]: bottom; null mean blank row
	private int[] getRowMinMax(Worksheet sheet, Row rowobj, int minc, int maxc) {
		if (rowobj == null) { //check if no cell at all!
			return null;
		}
		final int row = rowobj.getRowNum();
		int minr = row;
		int maxr = row;
		boolean allblank = true;

		//initial minc
		final int[] minrng = getCellMinMax(sheet, row, minc);
		if (minrng != null) {
			final int l = minrng[0];
			final int t = minrng[1];
			final int b = minrng[3];
			if (minr > t) minr = t;
			if (maxr < b) maxr = b;
			minc = l;
			allblank = false;
		}
		
		//initial maxc
		if (maxc > (minrng != null ? minrng[2] : minc)) {
			final int[] maxrng = getCellMinMax(sheet, row, maxc);
			if (maxrng != null) {
				final int t = maxrng[1];
				final int r = maxrng[2];
				final int b = maxrng[3];
				if (minr > t) minr = t;
				if (maxr < b) maxr = b;
				maxc = r;
				allblank = false;
			}
		} else if (minrng != null) {
			maxc = minrng[2];
		}
		
		final int lc = rowobj.getFirstCellNum();
		final int rc = rowobj.getLastCellNum() - 1;
		final int left = minc > 0 ? minc - 1 : 0;
		final int right = maxc + 1;

		//locate new minc to its left
		for(int c = left; c >= lc; --c) {
			final int[] rng = getCellMinMax(sheet, row, c);
			if (rng != null) {
				final int l = rng[0];
				final int t = rng[1];
				final int b = rng[3];
				minc = c = l;
				if (minr > t) minr = t;
				if (maxr < b) maxr = b;
				allblank = false;
			} else {
				break;
			}
		}
		//locate new maxc to its right
		for(int c = right; c <= rc; ++c) {
			final int[] rng = getCellMinMax(sheet, row, c);
			if (rng != null) {
				final int t = rng[1];
				final int r = rng[2];
				final int b = rng[3];
				maxc = c = r;
				if (minr > t) minr = t;
				if (maxr < b) maxr = b;
				allblank = false;
			} else {
				break;
			}
		}
		return allblank ? null : new int[] {minc, minr, maxc, maxr};
	}
	
	//given a cell return the maximum range
	private CellRangeAddress getCurrentRegion(Worksheet sheet, int row, int col) {
		int minc = col;
		int maxc = col;
		int minr = Integer.MAX_VALUE;
		int maxr = -1;
		final Row roworg = sheet.getRow(row);
		final int[] ltrb = getRowMinMax(sheet, roworg, minc, maxc);
		if (ltrb != null) {
			minc = ltrb[0];
			minr = ltrb[1];
			maxc = ltrb[2];
			maxr = ltrb[3];
		}
		
		int ru = row > 0 ? row - 1 : row;
		int rd = row + 1;
		
		boolean stopu = ru == row;
		boolean stopd = false;
		do {
			//for row above
			if (!stopu) {
				final Row rowu = sheet.getRow(ru);
				final int[] ltrbu = getRowMinMax(sheet, rowu, minc, maxc);
				if (ltrbu != null) {
					if (minc != ltrbu[0] || maxc != ltrbu[2]) {  //minc or maxc changed!
						stopd = false;
						minc = ltrbu[0];
						maxc = ltrbu[2];
					}
					if (minr > ltrbu[1]) {
						minr = ltrbu[1];
					}
					if (ru > 0) {
						--ru;
					} else {
						stopu = true; //no more row above!
					}
				} else { //blank row
					stopu = true;
				}
			}

			//for row below
			if (!stopd) {
				final Row rowd = sheet.getRow(rd);
				final int[] ltrbd = getRowMinMax(sheet, rowd, minc, maxc);
				if (ltrbd != null) {
					if (minc != ltrbd[0] || maxc != ltrbd[2]) { //minc and maxc changed
						stopu = false;
						minc = ltrbd[0];
						maxc = ltrbd[2];
					}
					if (maxr < ltrbd[3]) {
						maxr = ltrbd[3];
					}
					++rd;
				} else { //blank row
					stopd = true;
				}
			}
		} while(!stopu || !stopd);
		
		if (minr == Integer.MAX_VALUE && maxr < 0) { //all blanks in 9 cells!
			return null;
		}
		return new CellRangeAddress(minr, maxr, minc, maxc);
	}

	//TODO:
	private static final String ALL_BLANK_MSG = "Cannot find the range. Please select a cell within the range and try again!";
	@Override
	public AutoFilter autoFilter() {
		synchronized (_sheet) {
			CellRangeAddress affectedArea;
			if(_sheet.isAutoFilterMode()){
				affectedArea = _sheet.removeAutoFilter();
				Range unhideArea = Ranges.range(_sheet,
						affectedArea.getFirstRow(),affectedArea.getFirstColumn(),
						affectedArea.getLastRow(),affectedArea.getLastColumn());
				unhideArea.getRows().setHidden(false);
			} else {
				//The logic to decide the actual affected range to implement autofilter:
				//If it's a multi cell range, it's the range intersect with largest range of the sheet.
				//If it's a single cell range, it has to be extend to a continuous range by looking up the near 8 cells of the single cell.
				affectedArea = new CellRangeAddress(getRow(), getLastRow(), getColumn(), getLastColumn());
				if (BookHelper.isOneCell(_sheet, affectedArea)) { //only one cell selected(include merged one), try to look the max range surround by blank cells 
					CellRangeAddress maxRange = getCurrentRegion(_sheet, getRow(), getColumn());
					if (maxRange == null) {
						throw new RuntimeException(ALL_BLANK_MSG);
					}
					affectedArea = maxRange;
				} else {
					CellRangeAddress largeRange = getLargestRange(_sheet); //get the largest range that contains non-blank cells
					if (largeRange == null) {
						throw new RuntimeException(ALL_BLANK_MSG);
					}
					int left = largeRange.getFirstColumn();
					int top = largeRange.getFirstRow();
					int right = largeRange.getLastColumn();
					int bottom = largeRange.getLastRow();
					if (left < getColumn()) {
						left = getColumn();
					}
					if (right > getLastColumn()) {
						right = getLastColumn();
					}
					if (top < getRow()) {
						top = getRow();
					}
					if (bottom > getLastRow()) {
						bottom = getLastRow();
					}
					if (top > bottom || left > right) {
						throw new RuntimeException(ALL_BLANK_MSG);
					}
					affectedArea = new CellRangeAddress(top, bottom, left, right);
				}
				_sheet.setAutoFilter(affectedArea);
			}
			 
			//I have to know the top row area to show/remove the combo button
			//changed by this autofilter action, it's not the same area
			//and then send ON_CONTENTS_CHANGE event
			RangeImpl buttonChange = (RangeImpl) Ranges.range(_sheet, 
					affectedArea.getFirstRow(),affectedArea.getFirstColumn(),
					affectedArea.getFirstRow(),affectedArea.getLastColumn());
			
			BookHelper.notifyBtnChanges(new HashSet<Ref>(buttonChange.getRefs()));
			
			return _sheet.getAutoFilter();
		}
	}
	
	private boolean canUnhide(AutoFilter af, FilterColumn fc, int row, int col) {
		final Set<Ref> all = new HashSet<Ref>();
		final List<FilterColumn> fltcs = af.getFilterColumns();
		for(FilterColumn fltc: fltcs) {
			if (fc.equals(fltc)) continue;
			if (shallHide(fltc, row, col)) { //any FilterColumn that shall hide the row
				return false;
			}
		}
		return true;
	}
	
	private boolean shallHide(FilterColumn fc, int row, int col) {
		final Cell cell = BookHelper.getCell(_sheet, row, col + fc.getColId());
		final boolean blank = BookHelper.isBlankCell(cell); 
		final String val =  blank ? "=" : BookHelper.getCellText(cell); //"=" means blank!
		final Set critera1 = fc.getCriteria1();
		return critera1 != null && !critera1.isEmpty() && !critera1.contains(val);
	}

	@Override
	public void showAllData() {
		synchronized (_sheet) {
			AutoFilter af = _sheet.getAutoFilter();
			if (af == null) { //no AutoFilter to apply 
				return;
			}
			final CellRangeAddress afrng = af.getRangeAddress();
			final List<FilterColumn> fcs = af.getFilterColumns();
			if (fcs == null)
				return;
			for(FilterColumn fc : fcs) {
				BookHelper.setProperties(fc, null, AutoFilter.FILTEROP_VALUES, null, null); //clear all filter
			}
			final int row1 = afrng.getFirstRow();
			final int row = row1 + 1;
			final int row2 = afrng.getLastRow();
			final int col1 = afrng.getFirstColumn();
			final int col2 = afrng.getLastColumn();
			final Set<Ref> all = new HashSet<Ref>();
			for (int r = row; r <= row2; ++r) {
				final Row rowobj = _sheet.getRow(r);
				if (rowobj != null && rowobj.getZeroHeight()) { //a hidden row
					final int left = rowobj.getFirstCellNum();
					final int right = rowobj.getLastCellNum() - 1;
					final RangeImpl rng = (RangeImpl) new RangeImpl(r, left, r, right, _sheet, _sheet); 
					all.addAll(rng.getRefs());
					rng.getRows().setHidden(false); //unhide
				}
			}
	
			BookHelper.notifyCellChanges(_sheet.getBook(), all); //unhidden row must reevaluate
			
			//update button
			final RangeImpl buttonChange = (RangeImpl) Ranges.range(_sheet, row1, col1, row1, col2);
			BookHelper.notifyBtnChanges(new HashSet<Ref>(buttonChange.getRefs()));
		}
	}
	
	@Override
	public void applyFilter() {
		synchronized (_sheet) {
			AutoFilter af = _sheet.getAutoFilter();
			if (af == null) { //no AutoFilter to apply 
				return;
			}
			final CellRangeAddress affectedArea = af.getRangeAddress();
			final int row1 = affectedArea.getFirstRow();
			final int col1 = affectedArea.getFirstColumn(); 
			final int row = row1 + 1;
			final int row2 = affectedArea.getLastRow();
			final int col2 = affectedArea.getLastColumn();
			
			final Set<Ref> all = new HashSet<Ref>();
			for (int r = row; r <= row2; ++r) {
				boolean hidden = false;
				final List<FilterColumn> fcs = af.getFilterColumns();
				if (fcs == null)
					return;
				for(FilterColumn fc : fcs) {
					if (shallHide(fc, r, col1)) {
						hidden = true;
						break;
					}
				}
				if (hidden) { //to be hidden
					final Row rowobj = _sheet.getRow(r);
					if (rowobj == null || !rowobj.getZeroHeight()) { //a non-hidden row
						new RangeImpl(r, col1, _sheet, _sheet).getRows().setHidden(true);
					}
				} else { //to be shown
					final Row rowobj = _sheet.getRow(r);
					if (rowobj != null && rowobj.getZeroHeight()) { //a hidden row
						final int left = rowobj.getFirstCellNum();
						final int right = rowobj.getLastCellNum() - 1;
						final RangeImpl rng = (RangeImpl) new RangeImpl(r, left, r, right, _sheet, _sheet); 
						all.addAll(rng.getRefs());
						rng.getRows().setHidden(false); //unhide
					}
				}
			}
			BookHelper.notifyCellChanges(_sheet.getBook(), all); //unhidden row must reevaluate
			
			//update button
			final RangeImpl buttonChange = (RangeImpl) Ranges.range(_sheet, row1, col1, row1, col2);
			BookHelper.notifyBtnChanges(new HashSet<Ref>(buttonChange.getRefs()));
		}
	}

	@Override
	public AutoFilter autoFilter(int field, Object criteria1, int filterOp, Object criteria2, Boolean visibleDropDown) {
		synchronized (_sheet) {
			AutoFilter af = _sheet.getAutoFilter();
			if (af == null) {
				af = autoFilter();
			}
			final FilterColumn fc = BookHelper.getOrCreateFilterColumn(af, field-1);
			BookHelper.setProperties(fc, criteria1, filterOp, criteria2, visibleDropDown);
			
			//update rows
			final CellRangeAddress affectedArea = af.getRangeAddress();
			final int row1 = affectedArea.getFirstRow();
			final int col1 = affectedArea.getFirstColumn(); 
			final int col =  col1 + field - 1;
			final int row = row1 + 1;
			final int row2 = affectedArea.getLastRow();
			final Set cr1 = fc.getCriteria1();
	
			final Set<Ref> all = new HashSet<Ref>();
			for (int r = row; r <= row2; ++r) {
				final Cell cell = BookHelper.getCell(_sheet, r, col); 
				final String val = BookHelper.isBlankCell(cell) ? "=" : BookHelper.getCellText(cell); //"=" means blank!
				if (cr1 != null && !cr1.isEmpty() && !cr1.contains(val)) { //to be hidden
					final Row rowobj = _sheet.getRow(r);
					if (rowobj == null || !rowobj.getZeroHeight()) { //a non-hidden row
						new RangeImpl(r, col, _sheet, _sheet).getRows().setHidden(true);
					}
				} else { //candidate to be shown (other FieldColumn might still hide this row!
					final Row rowobj = _sheet.getRow(r);
					if (rowobj != null && rowobj.getZeroHeight() && canUnhide(af, fc, r, col1)) { //a hidden row and no other hidden filtering
						final int left = rowobj.getFirstCellNum();
						final int right = rowobj.getLastCellNum() - 1;
						final RangeImpl rng = (RangeImpl) new RangeImpl(r, left, r, right, _sheet, _sheet); 
						all.addAll(rng.getRefs());
						rng.getRows().setHidden(false); //unhide
					}
				}
			}
			
			BookHelper.notifyCellChanges(_sheet.getBook(), all); //unhidden row must reevaluate
			
			//update button
			final RangeImpl buttonChange = (RangeImpl) Ranges.range(_sheet, row1, col, row1, col);
			BookHelper.notifyBtnChanges(new HashSet<Ref>(buttonChange.getRefs()));
			
			return af;
		}
	}

	@Override
	public Chart addChart(ClientAnchor anchor, ChartData data, ChartType type,
			ChartGrouping grouping, LegendPosition pos) {
		synchronized (_sheet) {
			final DrawingManager dm = ((SheetCtrl)_sheet).getDrawingManager();
			final XSSFChartX chartX = (XSSFChartX) dm.addChartX(_sheet, anchor, data, type, grouping, pos);
			final RangeImpl rng = (RangeImpl) Ranges.range(_sheet, anchor.getRow1(), anchor.getCol1(), anchor.getRow2(), anchor.getCol2());
			final Collection<Ref> refs = rng.getRefs();
			if (refs != null && !refs.isEmpty()) {
				final Ref ref = refs.iterator().next();
				BookHelper.notifyChartAdd(ref, chartX);
			}
			return chartX.getChart();
		}
	}

	@Override
	public Picture addPicture(ClientAnchor anchor, byte[] image, int format) {
		synchronized (_sheet) {
			DrawingManager dm = ((SheetCtrl)_sheet).getDrawingManager();
			final Picture picture = dm.addPicture(_sheet, anchor, image, format);
			final RangeImpl rng = (RangeImpl) Ranges.range(_sheet, anchor.getRow1(), anchor.getCol1(), anchor.getRow2(), anchor.getCol2());
			final Collection<Ref> refs = rng.getRefs();
			if (refs != null && !refs.isEmpty()) {
				final Ref ref = refs.iterator().next();
				BookHelper.notifyPictureAdd(ref, picture);
			}
			return picture;
		}
	}
	
	@Override
	public void deletePicture(Picture picture) {
		synchronized (_sheet) {
			DrawingManager dm = ((SheetCtrl)_sheet).getDrawingManager();
			ClientAnchor anchor = picture.getPreferredSize();
			final RangeImpl rng = (RangeImpl) Ranges.range(_sheet, anchor.getRow1(), anchor.getCol1(), anchor.getRow2(), anchor.getCol2());
			final Collection<Ref> refs = rng.getRefs();
			dm.deletePicture(_sheet, picture); //must after getPreferredSize() or anchor is gone!
			if (refs != null && !refs.isEmpty()) {
				final Ref ref = refs.iterator().next();
				BookHelper.notifyPictureDelete(ref, picture);
			}
		}
	}

	@Override
	public void movePicture(Picture picture, ClientAnchor anchor) {
		synchronized (_sheet) {
			DrawingManager dm = ((SheetCtrl)_sheet).getDrawingManager();
			dm.movePicture(_sheet, picture, anchor);
			final RangeImpl rng = (RangeImpl) Ranges.range(_sheet, anchor.getRow1(), anchor.getCol1(), anchor.getRow2(), anchor.getCol2());
			final Collection<Ref> refs = rng.getRefs();
			if (refs != null && !refs.isEmpty()) {
				final Ref ref = refs.iterator().next();
				BookHelper.notifyPictureUpdate(ref, picture);
			}
		}
	}

	@Override
	public void moveChart(Chart chart, ClientAnchor anchor) {
		synchronized (_sheet) {
			DrawingManager dm = ((SheetCtrl)_sheet).getDrawingManager();
			dm.moveChart(_sheet, chart, anchor);
			final RangeImpl rng = (RangeImpl) Ranges.range(_sheet, anchor.getRow1(), anchor.getCol1(), anchor.getRow2(), anchor.getCol2());
			final Collection<Ref> refs = rng.getRefs();
			if (refs != null && !refs.isEmpty()) {
				final Ref ref = refs.iterator().next();
				BookHelper.notifyChartUpdate(ref, chart);
			}
		}
	}

	@Override
	public void deleteChart(Chart chart) {
		synchronized (_sheet) {
			DrawingManager dm = ((SheetCtrl)_sheet).getDrawingManager();
			ClientAnchor anchor = chart.getPreferredSize();
			final RangeImpl rng = (RangeImpl) Ranges.range(_sheet, anchor.getRow1(), anchor.getCol1(), anchor.getRow2(), anchor.getCol2());
			final Collection<Ref> refs = rng.getRefs();
			dm.deleteChart(_sheet, chart); //must after getPreferredSize() or anchor is gone!
			if (refs != null && !refs.isEmpty()) {
				final Ref ref = refs.iterator().next();
				BookHelper.notifyChartDelete(ref, chart);
			}
		}
	}

	@Override
	public void notifyMoveFriendFocus(Object token) {
		Ref ref = _refs != null && !_refs.isEmpty() ? _refs.iterator().next() : null;
		if (ref != null) {
			BookHelper.notifyMoveFriendFocus(ref, token);
		}
	}

	@Override
	public void notifyDeleteFriendFocus(Object token) {
		Ref ref = _refs != null && !_refs.isEmpty() ? _refs.iterator().next() : null;
		if (ref != null) {
			BookHelper.notifyDeleteFriendFocus(ref, token);
		}
	}

	@Override
	public void deleteSheet() {
		synchronized (_sheet.getBook()) {
			Ref ref = _refs != null && !_refs.isEmpty() ? _refs.iterator().next() : null;
			if (ref != null) {
				final Book book = _sheet.getBook();
				if (book != null) {
					final int index = book.getSheetIndex(_sheet);
					if (index != -1) {
						final int sheetCount = book.getNumberOfSheets();
						if (sheetCount == 1) {
							try {
								Messagebox.show("A workbook must contain at least one visible worksheet");
							} catch (InterruptedException e) {
							}
						}
						final String delSheetName = _sheet.getSheetName(); //must getName before remove
						book.removeSheetAt(index);
						final int newIndex =  index < (sheetCount - 1) ? index : (index - 1);
						final String newSheetName = book.getSheetName(newIndex);
						BookHelper.notifyDeleteSheet(ref, new Object[] {delSheetName, newSheetName});
					}
				}
			}
		}
	}

	@Override
	public boolean isCustomHeight() {
		Ref ref = _refs != null && !_refs.isEmpty() ? _refs.iterator().next() : null;
		if (ref != null) {
			final int tRow = ref.getTopRow();
			final RefSheet refSheet = ref.getOwnerSheet();
			final Row row = getRow(tRow, refSheet);
			if (row != null) {
				return row.isCustomHeight();
			}
		}
		return false;
	}
}
