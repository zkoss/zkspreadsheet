/* RangeImpl.java

	Purpose:
		
	Description:
		
	History:
		Mar 10, 2010 2:54:55 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.sys.impl;

import java.net.URL;
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

import org.zkoss.lang.Strings;
import org.zkoss.poi.hssf.usermodel.HSSFSheet;
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
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.poi.ss.usermodel.ZssChartX;
import org.zkoss.poi.ss.usermodel.charts.ChartData;
import org.zkoss.poi.ss.usermodel.charts.ChartGrouping;
import org.zkoss.poi.ss.usermodel.charts.ChartType;
import org.zkoss.poi.ss.usermodel.charts.LegendPosition;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.poi.xssf.usermodel.XSSFChartX;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zss.api.IllegalOpArgumentException;
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
import org.zkoss.zss.model.sys.XAreas;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XFormatText;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.XRanges;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zul.Messagebox;
/**
 * Implementation of {@link XRange} which plays a facade to operate on the spreadsheet data models
 * and maintain the reference dependency. 
 * @author henrichen
 *
 */
public class XRangeImpl implements XRange {
	private final XSheet _sheet;
	private int _left = Integer.MAX_VALUE;
	private int _top = Integer.MAX_VALUE;
	private int _right = Integer.MIN_VALUE;
	private int _bottom = Integer.MIN_VALUE;
	private Set<Ref> _refs;

	public XRangeImpl(int row, int col, XSheet sheet, XSheet lastSheet) {
		_sheet = sheet;
		initCellRef(row, col, sheet, lastSheet);
	}
	
	private void initCellRef(int row, int col, XSheet sheet, XSheet lastSheet) {
		final XBook book  = (XBook) sheet.getWorkbook();
		final int s1 = book.getSheetIndex(sheet);
		final int s2 = book.getSheetIndex(lastSheet);
		final int sb = Math.min(s1,s2);
		final int se = Math.max(s1,s2);
		for (int s = sb; s <= se; ++s) {
			final XSheet sht = book.getWorksheetAt(s);
			final RefSheet refSht = BookHelper.getOrCreateRefBook(book).getOrCreateRefSheet(sht.getSheetName());
			addRef(new CellRefImpl(row, col, refSht));
		}
	}
	
	public XRangeImpl(int tRow, int lCol, int bRow, int rCol, XSheet sheet, XSheet lastSheet) {
		_sheet = sheet;
		if (tRow == bRow && lCol == rCol) 
			initCellRef(tRow, lCol, sheet, lastSheet);
		else
			initAreaRef(tRow, lCol, bRow, rCol, sheet, lastSheet);
	}
	
	/*package*/ XRangeImpl(Ref ref, XSheet sheet) {
		_sheet = BookHelper.getSheet(sheet, ref.getOwnerSheet());
		addRef(ref);
	}
	
	private XRangeImpl(Set<Ref> refs, XSheet sheet) {
		_sheet = sheet;
		for(Ref ref : refs) {
			addRef(ref);
		}
	}
	
	private void initAreaRef(int tRow, int lCol, int bRow, int rCol, XSheet sheet, XSheet lastSheet) {
		final XBook book  = (XBook) sheet.getWorkbook();
		final int s1 = book.getSheetIndex(sheet);
		final int s2 = book.getSheetIndex(lastSheet);
		final int sb = Math.min(s1,s2);
		final int se = Math.max(s1,s2);
		for (int s = sb; s <= se; ++s) {
			final XSheet sht = book.getWorksheetAt(s);
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
	public XSheet getSheet() {
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

	private int getHyperlinkType(String address) {
		if (address != null && !isDirectHyperlink()) {
			final String addr = address.toLowerCase(); // ZSS-288: support more scheme according to POI code, see  org.zkoss.poi.ss.formula.functions.Hyperlink
			if (addr.startsWith("http://") || addr.startsWith("https://")) {
				return org.zkoss.poi.ss.usermodel.Hyperlink.LINK_URL;
			} else if (addr.startsWith("mailto:")) {
				return org.zkoss.poi.ss.usermodel.Hyperlink.LINK_EMAIL;
			} // ZSS-288: don't support auto-create hyperlink for DOCUMENT and FILE type
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
	public XFormatText getFormatText() {
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
				final XSheet sheet = BookHelper.getSheet(_sheet, refSheet);
				
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
				
				//ZSS-414 get NPE when paste from client text
				//in bookhelper, public static Set<Ref>[] setCellValue(Cell cell, String value) {
				//it ignore to set a new create cell value (but type String), cause this error in 2003
				//however, we should ignore to create a cell if text is empty
				if(cell==null && (txt==null || txt.length()==0)){
					return;
				}
				
				final Object[] values = BookHelper.editTextToValue(txt, cell);
				final XSheet sheet = BookHelper.getSheet(_sheet, refSheet);
				
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
	
	@Override
	public void notifyChange() {
		synchronized (_sheet) {
			if (_refs != null && !_refs.isEmpty()) {
				final XBook book = (XBook) _sheet.getWorkbook();
				BookHelper.notifyCellChanges(book, _refs);
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
			final XSheet sheet = BookHelper.getSheet(_sheet, refSheet);
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
			final XBook book  = (XBook) _sheet.getWorkbook();
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
		final XBook book = BookHelper.getBook(_sheet, refSheet);
		if (book != null) {
			final XSheet sheet = book.getWorksheet(refSheet.getSheetName());
			if (sheet != null) {
				return sheet.getRow(rowIndex);
			}
		}
		return null;
	}
	
	private Cell getCell(int rowIndex, int colIndex, RefSheet refSheet) {
		//locate the model book and sheet of the refSheet
		final XBook book = BookHelper.getBook(_sheet, refSheet);
		if (book != null) {
			final XSheet sheet = book.getWorksheet(refSheet.getSheetName());
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
		final XBook book = BookHelper.getBook(_sheet, refSheet);
		final String sheetname = refSheet.getSheetName();
		final XSheet sheet = getOrCreateSheet(book, sheetname);
		final Row row = getOrCreateRow(sheet, rowIndex);
		return getOrCreateCell(row, colIndex, cellType);
	}
	
	private XSheet getOrCreateSheet(XBook book, String sheetName) {
		XSheet sheet = book.getWorksheet(sheetName);
		if (sheet == null) {
			sheet = (XSheet)book.createSheet(sheetName);
		}
		return sheet;
	}
	
	private Row getOrCreateRow(XSheet sheet, int rowIndex) {
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
				final XSheet sheet = BookHelper.getSheet(_sheet, refSheet);
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
				final XSheet sheet = BookHelper.getSheet(_sheet, refSheet);
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
		//must delete and add in batch, or merge ranges can interfere to each other
		for(MergeChange change : info.getMergeChanges()) {
			final Ref orgMerge = change.getOrgMerge();
			if (orgMerge != null) {
				refBook.publish(new SSDataEvent(SSDataEvent.ON_MERGE_DELETE, orgMerge, orient));
			}
		}
		// ZSS-354: delete merge before cell update, because of this merge range is based on original cell range
		if (event != null && ref != null) {
			refBook.publish(new SSDataEvent(event, ref, orient));
		}
		for(MergeChange change : info.getMergeChanges()) {
			final Ref merge = change.getMerge();
			if (merge != null) {
				refBook.publish(new SSDataEvent(SSDataEvent.ON_MERGE_ADD, merge, orient));
			}
		}
		BookHelper.reevaluateAndNotify((XBook) _sheet.getWorkbook(), last, all);
	}

	@Override
	public XRange pasteSpecial(int pasteType, int operation, boolean SkipBlanks, boolean transpose) {
		//TODO
		//clipboardRange.pasteSpecial(this, pasteType, pasteOp, skipBlanks, transpose);
		return null;
	}
	
	//enforce lock sequence
	private XSheet[] getLockSheets(XRange dstRange) {
		final int srcIndex = _sheet.getBook().getSheetIndex(_sheet);
		final XSheet dstSheet = dstRange.getSheet();
		final int dstIndex = dstSheet.getBook().getSheetIndex(dstSheet);
		final XSheet[] sheets = new XSheet[2];
		final XSheet sheet1 = srcIndex > dstIndex ? _sheet : dstSheet;
		final XSheet sheet2 = srcIndex > dstIndex ? dstSheet : _sheet;
		return new XSheet[] {sheet1, sheet2};
	}
	
	@Override
	public XRange pasteSpecial(XRange dstRange, int pasteType, int pasteOp, boolean skipBlanks, boolean transpose) {
		final XSheet[] sheets = getLockSheets(dstRange); //enforce lock sequence
		synchronized(sheets[0]) { 
			synchronized(sheets[1]) {
				final Ref ref = paste0(dstRange, pasteType, pasteOp, skipBlanks, transpose);
				return ref == null ? null : new XRangeImpl(ref, BookHelper.getSheet(_sheet, ref.getOwnerSheet()));
			}
		}
	}
	
	@SuppressWarnings("unchecked")
	@Override
	public void sort(XRange rng1, boolean desc1, XRange rng2, int type, boolean desc2, XRange rng3, boolean desc3, int header, int orderCustom,
			boolean matchCase, boolean sortByRows, int sortMethod, int dataOption1, int dataOption2, int dataOption3) {
		synchronized(_sheet) {
			final Ref key1 = rng1 != null ? ((XRangeImpl)rng1).getRefs().iterator().next() : null;
			final Ref key2 = rng2 != null ? ((XRangeImpl)rng2).getRefs().iterator().next() : null;
			final Ref key3 = rng3 != null ? ((XRangeImpl)rng3).getRefs().iterator().next() : null;
			if (_refs != null && !_refs.isEmpty()) {
				final Ref ref = _refs.iterator().next();
				final int tRow = ref.getTopRow();
				final int lCol = ref.getLeftCol();
				final int bRow = ref.getBottomRow();
				final int rCol = ref.getRightCol();
				final XSheet sheet = BookHelper.getSheet(_sheet, ref.getOwnerSheet());
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
	public XRange copy(XRange dstRange) {
		return copy(dstRange,false);
	}
	
	@Override
	public XRange copy(XRange dstRange, boolean cut) {
		final XSheet[] sheets = getLockSheets(dstRange); //enforce lock sequence
		synchronized(sheets[0]) { //enforce lock sequence
			synchronized(sheets[1]) {
				XRange pasteRange = copy0(dstRange);
				if(!cut) {
					return pasteRange;
				}
				
				Ref srcRef = _refs.iterator().next();
				final Ref dstRef = ((XRangeImpl)pasteRange).getRefs().iterator().next();
				OverlapState state = getOverlapState(dstRef, srcRef);
				
				if(state == OverlapState.INCLUDED_OUTSIDE || state == OverlapState.INCLUDED_INSIDE) {
					return pasteRange;
				}
				
				List<Ref> areas = removeIntersect(dstRef, srcRef);
				if(areas.isEmpty()) { // no overlap, clear all source
					clearContents();
					setStyle(null);
					unMerge();
				} else { // clear area contents
					for(Ref area : areas) {
						XRange range = new XRangeImpl(area, BookHelper.getSheet(_sheet, area.getOwnerSheet()));
						range.clearContents();
						range.setStyle(null);
						range.unMerge();
					}
				}
				return pasteRange;
			}
		}
	}

	private XRange copy0(XRange dstRange) {
		final Ref ref = paste0(dstRange, XRange.PASTE_ALL, XRange.PASTEOP_NONE, false, false);
		return ref == null ? null : new XRangeImpl(ref, BookHelper.getSheet(_sheet, ref.getOwnerSheet()));
	}
	
	private Ref paste0(XRange dstRange, int pasteType, int pasteOp, boolean skipBlanks, boolean transpose) {
		if (_refs != null && !_refs.isEmpty() && !((XRangeImpl)dstRange).getRefs().isEmpty()) {
			//check if follow the copy rule
			//destination range allow only one contiguous reference
			if (((XRangeImpl)dstRange).getRefs().size() > 1) {
				throw new IllegalOpArgumentException("Command cannot be used on multiple selections");
			}
			//source range can handle only same rows/columns multiple references
			Iterator<Ref> it = _refs.iterator();
			Ref ref1 = it.next();
			int srcRowCount = ref1.getRowCount();
			int srcColCount = ref1.getColumnCount();
			final Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();	// destination reference
			
			if(transpose) { // transpose paste on overlap area is a illegal operation
				final int dstRightCol = srcColCount > dstRef.getColumnCount() ? dstRef.getLeftCol() + srcColCount - 1 : dstRef.getRightCol();
				final int dstBottomRow = srcRowCount > dstRef.getRowCount() ? dstRef.getBottomRow() + srcRowCount - 1 : dstRef.getBottomRow();
				final Ref extendedDstRef = ((XRangeImpl) XRanges.range(dstRange.getSheet(), dstRef.getTopRow(), dstRef.getLeftCol(), dstBottomRow, dstRightCol)).getRefs().iterator().next();
				if(getOverlapState(ref1, extendedDstRef) != OverlapState.NOT_OVERLAP) {
					throw new IllegalOpArgumentException("Cannot transpose paste to overlapped range");
				}
			}
			
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
							throw new IllegalOpArgumentException("Command cannot be used on multiple selections");
						}
						if (srcRefs.isEmpty()) {
							srcRefs.put(new Integer(tRow), ref1); //sorted on Row
						}
						srcRefs.put(new Integer(ref.getTopRow()), ref);
						sameCol = true;
						srcRowCount += ref.getRowCount();
					} else if (tRow == ref.getTopRow() && bRow == ref.getBottomRow()) { //same row
						if (sameCol) { //cannot be both sameRow and sameColumn
							throw new IllegalOpArgumentException("Command cannot be used on multiple selections");
						}
						if (srcRefs.isEmpty()) {
							srcRefs.put(Integer.valueOf(lCol), ref1); //sorted on column
						}
						srcRefs.put(Integer.valueOf(ref.getLeftCol()), ref);
						sameRow = true;
						srcColCount += ref.getColumnCount();
					} else { //not the same column or same row
						throw new IllegalOpArgumentException("Command cannot be used on multiple selections");
					}
				}
				pasteType = pasteType + XRange.PASTE_VALUES; //no formula 
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
	
	public static enum OverlapState {
		NOT_OVERLAP,ON_RIGHT_BOTTOM, ON_LEFT_BOTTOM, ON_LEFT_TOP, ON_RIGHT_TOP, INCLUDED_INSIDE, INCLUDED_OUTSIDE,
		ON_TOP, ON_BOTTOM, ON_RIGHT, ON_LEFT;
	}
	
	// is cell (row,col) is inside the ref area?
	private boolean isInside(Ref ref, int row, int col) {
		return row >= ref.getTopRow() && row <= ref.getBottomRow() && col >= ref.getLeftCol() && col <= ref.getRightCol();
	}

	public OverlapState getOverlapState(Ref ref1, Ref ref2) {
		
		if(ref1.getOwnerSheet() != ref2.getOwnerSheet()) {
			return OverlapState.NOT_OVERLAP; // not the same sheet
		}
		
		final int ref1_Height = ref1.getRowCount();
		final int ref1_Width = ref1.getColumnCount();
		final int ref2_Height = ref2.getRowCount();
		final int ref2_Width = ref2.getColumnCount();
		
		// if ref2 is bigger than ref1, pass ref argument reversely.
		if(ref2_Height > ref1_Height && ref2_Width > ref1_Width) {
			if(getOverlapState(ref2, ref1) == OverlapState.INCLUDED_INSIDE) {
				return OverlapState.INCLUDED_OUTSIDE;
			}
		} else {
			
			final boolean topLeftInRef = isInside(ref1, ref2.getTopRow(), ref2.getLeftCol());
			final boolean topRightInRef = isInside(ref1, ref2.getTopRow(), ref2.getRightCol());
			final boolean bottomLeftInRef = isInside(ref1, ref2.getBottomRow(), ref2.getLeftCol());
			final boolean bottomRightInRef = isInside(ref1, ref2.getBottomRow(), ref2.getRightCol());
			
			// if the diagonal are inside, it will be the included case.
			if(topLeftInRef && bottomRightInRef) {
				return OverlapState.INCLUDED_INSIDE;
			}
			
			// if two points inside, they will be first four case.
			// then only one point inside, they will be last four case. 
			if(bottomRightInRef && topRightInRef) {
				return OverlapState.ON_LEFT; 
			} else if(bottomLeftInRef && topLeftInRef) {
				return OverlapState.ON_RIGHT;
			} else if(topRightInRef && topLeftInRef) {
				return OverlapState.ON_BOTTOM;
			} else if(bottomRightInRef && bottomLeftInRef) {
				return OverlapState.ON_TOP;
			} else if (topLeftInRef) {
				return OverlapState.ON_RIGHT_BOTTOM; // right bottom
			} else if(topRightInRef) {
				return OverlapState.ON_LEFT_BOTTOM; // left bottom
			} else if(bottomLeftInRef) {
				return OverlapState.ON_RIGHT_TOP; // right top
			} else if(bottomRightInRef) {
				return OverlapState.ON_LEFT_TOP; // left top
			}
		}
		
		return OverlapState.NOT_OVERLAP; // default is not overlap
		
	}
	
	/**
	 * get the intersection of a & b
	 * assume a and b are in the same sheet
	 * @return a intersection ref area (null if not overlap or unsupported case)
	 */
	public Ref intersect(Ref a, Ref b) {
		
		// get overlap state
		final OverlapState result = getOverlapState(a, b);
		
		final int aLeftCol = a.getLeftCol();
		final int aRightCol = a.getRightCol();
		final int aTopRow = a.getTopRow();
		final int aBottomRow = a.getBottomRow();
		
		final int bLeftCol = b.getLeftCol();
		final int bRightCol = b.getRightCol();
		final int bTopRow = b.getTopRow();
		final int bBottomRow = b.getBottomRow();
		
		switch(result) {
			case ON_RIGHT_BOTTOM: // right bottom
				return new AreaRefImpl(bTopRow, bLeftCol, aBottomRow, aRightCol, a.getOwnerSheet());
			case ON_LEFT_BOTTOM: // left bottom
				return new AreaRefImpl(bTopRow, aLeftCol, aBottomRow, bRightCol, a.getOwnerSheet());
			case ON_LEFT_TOP: // left top
				return new AreaRefImpl(aTopRow, aLeftCol, bBottomRow, bRightCol, a.getOwnerSheet());
			case ON_RIGHT_TOP: // right top
				return new AreaRefImpl(aTopRow, bLeftCol, bBottomRow, aRightCol, a.getOwnerSheet());
			case INCLUDED_INSIDE: // ref2 inside ref1
				return b;
			case INCLUDED_OUTSIDE: // ref1 inside ref2
				return a;
			default: // not overlap, right, bottom, left..other case
				return null;
		}
	}
	
	/**
	 * remove "a intersect b" from "b" and split "b" into areas list  .
	 * assume a and b are in the same sheet.
	 * ignore "included" case.
	 * @return a list of complement areas
	 */
	public List<Ref> removeIntersect(Ref a, Ref b) {
		
		final OverlapState result = getOverlapState(a, b);
		
		List<Ref> resultSet = new ArrayList<Ref>(); // add ref in counter-clockwise
		
		if(result == OverlapState.NOT_OVERLAP) {
			return resultSet;
		}

		final int aLeftCol = a.getLeftCol();
		final int aRightCol = a.getRightCol();
		final int aTopRow = a.getTopRow();
		final int aBottomRow = a.getBottomRow();
		
		final int bLeftCol = b.getLeftCol();
		final int bRightCol = b.getRightCol();
		final int bTopRow = b.getTopRow();
		final int bBottomRow = b.getBottomRow();
		
		switch(result) {
			case ON_RIGHT_BOTTOM: // right bottom
				resultSet.add(new AreaRefImpl(bTopRow, aRightCol+1, aBottomRow, bRightCol, b.getOwnerSheet())); // Q1
				resultSet.add(new AreaRefImpl(aBottomRow+1, bLeftCol, bBottomRow, aRightCol, b.getOwnerSheet())); // Q3
				resultSet.add(new AreaRefImpl(aBottomRow+1, aRightCol+1, bBottomRow, bRightCol, b.getOwnerSheet())); // Q4
				break;
			case ON_LEFT_BOTTOM: // left bottom
				resultSet.add(new AreaRefImpl(bTopRow, bLeftCol, aBottomRow, aLeftCol-1, b.getOwnerSheet())); // Q2
				resultSet.add(new AreaRefImpl(aBottomRow+1, bLeftCol, bBottomRow, aLeftCol-1, b.getOwnerSheet())); // Q3 
				resultSet.add(new AreaRefImpl(aBottomRow+1, aLeftCol, bBottomRow, bRightCol, b.getOwnerSheet())); // Q4
				break;
			case ON_LEFT_TOP: // left top
				resultSet.add(new AreaRefImpl(bTopRow, aLeftCol, aTopRow-1, bRightCol, b.getOwnerSheet())); // Q1
				resultSet.add(new AreaRefImpl(bTopRow, bLeftCol, aTopRow-1, aLeftCol-1, b.getOwnerSheet())); // Q2
				resultSet.add(new AreaRefImpl(aTopRow, bLeftCol, bBottomRow, aLeftCol-1, b.getOwnerSheet())); // Q3
				break;
			case ON_RIGHT_TOP: // right top
				resultSet.add(new AreaRefImpl(bTopRow, bRightCol, aTopRow-1, bRightCol, b.getOwnerSheet())); // Q1
				resultSet.add(new AreaRefImpl(bTopRow, bLeftCol, aTopRow-1, aRightCol, b.getOwnerSheet())); // Q2
				resultSet.add(new AreaRefImpl(aTopRow, aRightCol+1, bBottomRow, bRightCol, b.getOwnerSheet())); // Q4
				break;
			case ON_RIGHT:
				resultSet.add(new AreaRefImpl(bTopRow, aRightCol+1, bBottomRow, bRightCol, b.getOwnerSheet()));
				break;
			case ON_LEFT:
				resultSet.add(new AreaRefImpl(bTopRow, bLeftCol, bBottomRow, aLeftCol-1, b.getOwnerSheet()));
				break;
			case ON_TOP:
				resultSet.add(new AreaRefImpl(bTopRow, bLeftCol, aTopRow-1, bRightCol, b.getOwnerSheet()));
				break;
			case ON_BOTTOM:
				resultSet.add(new AreaRefImpl(aBottomRow+1, bLeftCol, bBottomRow, bRightCol, b.getOwnerSheet()));
				break;
			default:
				break; // do nothing
		}
		
		return resultSet;
	}	

	@Override
	public void borderAround(BorderStyle lineStyle, String color) {
		setBorders(BookHelper.BORDER_OUTLINE, lineStyle, color);
	}
	
	@Override
	public void merge(boolean across) {
		// ZSS-290: unmerge before merging, or there will be multiple merged cell at same range
		unMerge();
		
		synchronized (_sheet) {
			if (_refs != null && !_refs.isEmpty()) {
				final Ref ref = _refs.iterator().next();
				final RefSheet refSheet = ref.getOwnerSheet();
				final RefBook refBook = refSheet.getOwnerBook();
				final int tRow = ref.getTopRow();
				final int lCol = ref.getLeftCol();
				final int bRow = ref.getBottomRow();
				final int rCol = ref.getRightCol();
				final XSheet sheet = BookHelper.getSheet(_sheet, refSheet);
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
				final XSheet sheet = BookHelper.getSheet(_sheet, refSheet);
				final ChangeInfo info = BookHelper.unMerge(sheet, tRow, lCol, bRow, rCol,true);
				notifyMergeChange(refBook, info, ref, SSDataEvent.ON_CONTENTS_CHANGE, SSDataEvent.MOVE_NO);
			}
		}
	}
	
	@Override
	public XRange getCells(int row, int col) {
		synchronized (_sheet) {
			final Ref ref = getRefs().iterator().next();
			final int col1 = ref.getLeftCol() + col;
			final int row1 = ref.getTopRow() + row;
			return new XRangeImpl(row1, col1, _sheet, _sheet);
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
				final XSheet sheet = BookHelper.getSheet(_sheet, refSheet);
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
		final XSheet dstSheet = BookHelper.getSheet(_sheet, dstRefSheet);
		final Ref pasteRef = new AreaRefImpl(dstT, dstL, dstB, dstR, dstRefSheet);
		return new XRangeImpl(pasteRef, dstSheet).isAnyCellProtected() ? null : pasteRef;
	}
	private Ref copyRefs(boolean sameRow, SortedMap<Integer, Ref> srcRefs, int srcColCount, int srcRowCount, int colRepeat, int rowRepeat, Ref dstRef, int pasteType, int pasteOp, boolean skipBlanks, boolean transpose, ChangeInfo info) {
		final Ref pasteRef = getPasteRef(srcRowCount, srcColCount, rowRepeat, colRepeat, dstRef, transpose);
		if (pasteRef == null) {
			return null;
		}
		if (pasteType == XRange.PASTE_COLUMN_WIDTHS) {
			final Integer lastKey = srcRefs.lastKey();
			final Ref srcRef = srcRefs.get(lastKey);
			final int srclCol = srcRef.getLeftCol();
			final XSheet srcSheet = BookHelper.getSheet(_sheet, srcRef.getOwnerSheet());
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
	
	private Ref copy(Ref srcRef, int srcColCount, int srcRowCount, XRange dstRange, Ref dstRef, int pasteType, int pasteOp, boolean skipBlanks, boolean transpose, ChangeInfo info) {
		if (pasteType == XRange.PASTE_COLUMN_WIDTHS) { //ignore transpose in such case when only one srcRef
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

	private void copyColumnWidths(XSheet srcSheet, int widthRepeatCount, int srclCol, Ref dstRef) {
		final int dstlCol = dstRef.getLeftCol();
		final int dstColCount = dstRef.getColumnCount();
		final RefSheet dstRefSheet = dstRef.getOwnerSheet();
		final XSheet dstSheet = BookHelper.getSheet(_sheet, dstRefSheet);
		for (int count = 0; count < dstColCount; ++count) {
			final int dstCol = dstlCol + count;
			final int srcCol = srclCol + count % widthRepeatCount; 
			final int char256 = srcSheet.getColumnWidth(srcCol);
			BookHelper.setColumnWidth(dstSheet, dstCol, dstCol, char256);
		}
		final XBook book = (XBook) dstSheet.getWorkbook();
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
		if (pasteType == XRange.PASTE_COLUMN_WIDTHS) {
			final int srclCol = srcRef.getLeftCol();
			final XSheet srcSheet = BookHelper.getSheet(_sheet, srcRef.getOwnerSheet());
			copyColumnWidths(srcSheet, srcColCount, srclCol, pasteRef);
			return pasteRef;
		}
		final int tRow = srcRef.getTopRow();
		final int lCol = srcRef.getLeftCol();
		final int bRow = srcRef.getBottomRow();
		final int rCol = srcRef.getRightCol();
		final int dstTopRow = dstRef.getTopRow();
		final int dstBottomRow = dstRef.getBottomRow();
		final int dstLeftCol = dstRef.getLeftCol();
		final int dstRightCol = dstRef.getRightCol();
		final RefSheet dstRefsheet = dstRef.getOwnerSheet();
		XSheet srcSheet = BookHelper.getSheet(_sheet, srcRef.getOwnerSheet()); // source may change when overlap issue!
		final XSheet dstSheet = BookHelper.getSheet(_sheet, dstRefsheet);
		final Set<Ref> toEval = info.getToEval();
		final Set<Ref> affected = info.getAffected();
		final List<MergeChange> mergeChanges = info.getMergeChanges();
		
		/**
		 * ZSS-315
		 * handle a special case: source is a single cell and destination is a single merged cell, 
		 * copy source to destination then merge it.
		 */
		if(srcRef.getRowCount() == 1 && srcRef.getColumnCount() == 1) { // source is a single cell
			final CellRangeAddress dstaddr = ((SheetCtrl)dstSheet).getMerged(dstTopRow, dstLeftCol);
			// Check whether destination is a single merged cell?
			if(dstaddr != null) {
				// copy source to merged cell
				// prevent null pointer, use getOrCreateCell helper.
				final Cell cell = BookHelper.getOrCreateCell(srcSheet, srcRef.getTopRow(), srcRef.getLeftCol()); // retrieve cell
				final ChangeInfo changeInfo0 = BookHelper.copyCell(cell, dstSheet, dstTopRow, dstLeftCol, pasteType, pasteOp, transpose);
				BookHelper.assignChangeInfo(toEval, affected, mergeChanges, changeInfo0);
				// merge cell (because cell always unmerge in BookHelper.copyCell)
				final ChangeInfo changeInfo1 = BookHelper.merge(dstSheet, dstTopRow, dstLeftCol, dstaddr.getLastRow(), dstaddr.getLastColumn(), false);
				BookHelper.assignChangeInfo(toEval, affected, mergeChanges, changeInfo1);
				return pasteRef;
			}
		}
		
		if (!transpose) {
			/**
			 * Issue: ZSS-300
			 * Problem: Overlapping copy get wrong result
			 * Root Cause: Copy dirty from source to destination
			 */
			int[][] repeatArea = new int[rowRepeat*colRepeat][2]; // [areaCount] => [TopRow][LeftCol]
			
			// calculate repeat area position
			for(int rr = rowRepeat, areaCount = 0; rr > 0; --rr) {
				for (int cr = colRepeat; cr > 0; --cr, ++areaCount) {
					repeatArea[areaCount][0] = dstRef.getTopRow() + (srcRef.getRowCount() * (rowRepeat - rr));
					repeatArea[areaCount][1] = dstRef.getLeftCol() + (srcRef.getColumnCount() * (colRepeat - cr));
				}
			}

			Set<String> blankMap = new HashSet<String>(); // blank mapping
			
			// Handle repeat area
			for(int areaCount = 0; areaCount < repeatArea.length; areaCount++) {
				
				// Decide start position of pointers
				// Default case
				// reference the source to first pasted area
				int srcRowPointer = areaCount <= 0 ? srcRef.getTopRow() : repeatArea[0][0];
				int srcColPointer = areaCount <= 0 ? srcRef.getLeftCol() : repeatArea[0][1];
				int dstRowPointer = areaCount <= 0 ? dstRef.getTopRow() : repeatArea[areaCount][0];
				int dstColPointer = areaCount <= 0 ? dstRef.getLeftCol() : repeatArea[areaCount][1];
				srcSheet = areaCount <= 0 ? srcSheet : dstSheet; // sheet should be replace too
				int rowStep = 1;
				int colStep = 1;
				
				// left corner of source
				int srcTopRow = srcRowPointer;
				int srcLeftCol = srcColPointer;
				
				if(srcRowPointer < dstRowPointer) { // row direction, destination is below the source
					srcRowPointer += srcRowCount - 1;
					dstRowPointer += srcRowCount - 1;
					rowStep = -1;
				}
				if(srcColPointer < dstColPointer) { // column direction, destination is on the right side of source
					srcColPointer += srcColCount - 1;
					dstColPointer += srcColCount - 1;
					colStep = -1;
				}
				
				for(int rowCount = srcRef.getRowCount(); rowCount > 0; rowCount--, srcRowPointer += rowStep, dstRowPointer += rowStep) { // Go through row
					
					for(int colCount = srcRef.getColumnCount(); colCount > 0; colCount--, srcColPointer += colStep, dstColPointer += colStep) { // Go through column
						
						final Cell cell = BookHelper.getCell(srcSheet, srcRowPointer, srcColPointer); // retrieve cell
						String hitBlankKey = Math.abs(srcRowPointer - srcTopRow) + "," + Math.abs(srcColPointer - srcLeftCol);
						boolean hitBlank = (blankMap.contains(hitBlankKey));
					
						// origin cell is not blank and current source cell is not null
						// although the current source cell is not a null cell, but it may be a blank cell
						// so we must check the skipBlanks configuration and whether the cell is blank
						if (!hitBlank && cell != null) {
							// if not skip blanks, we need to copy blank too!
							// if check whether the cell is blank			
							if(!(skipBlanks && cell.getCellType() == Cell.CELL_TYPE_BLANK)) { 
								final ChangeInfo changeInfo0 = BookHelper.copyCell(cell, dstSheet, dstRowPointer, dstColPointer, pasteType, pasteOp, transpose);
								BookHelper.assignChangeInfo(toEval, affected, mergeChanges, changeInfo0);
							}
						} else if (!skipBlanks) { // map hit blank or source cell is null, clear the destination
							final Set<Ref>[] refs = BookHelper.removeCell(dstSheet, dstRowPointer, dstColPointer);
							BookHelper.assignRefs(toEval, affected, refs);
						}
						
						// update blank map
						if(!hitBlank && cell != null && cell.getCellType() == Cell.CELL_TYPE_BLANK) { // cell is blank
							blankMap.add(hitBlankKey);
						}
						
						
					} // End go through column
					
					srcColPointer -= (colStep * srcColCount); // reset column pointer to origin
					dstColPointer -= (colStep * srcColCount); // pointer add (colStep * srcColCount) times so minus them back
					
				} // End go through row
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
				final XSheet sheet = BookHelper.getSheet(_sheet, refSheet);
				final Set<Ref> all = BookHelper.setBorders(sheet, tRow, lCol, bRow, rCol, borderIndex, lineStyle, color);
				if (all != null) {
					final XBook book = (XBook) _sheet.getWorkbook();
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
				final XSheet sheet = BookHelper.getSheet(_sheet, refSheet);
				final Set<Ref> all = BookHelper.setColumnWidth(sheet, lCol, rCol, char256);
				if (all != null) {
					final XBook book = (XBook) _sheet.getWorkbook();
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
				final XSheet sheet = BookHelper.getSheet(_sheet, refSheet);
				final Set<Ref> all = BookHelper.setRowHeight(sheet, tRow, bRow, (short) (points * 20), customHeight); //in twips, set customHeight
				if (all != null) {
					final XBook book = (XBook) _sheet.getWorkbook();
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
				final XSheet sheet = BookHelper.getSheet(_sheet, refSheet);
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
					final XSheet sheet = BookHelper.getSheet(_sheet, refSheet);
					final Set<Ref> refs = BookHelper.setCellStyle(sheet, tRow, lCol, bRow, rCol, style);
					all.addAll(refs);
				}
				if (!all.isEmpty()) {
					final XBook book = (XBook) _sheet.getWorkbook();
					BookHelper.notifyCellChanges(book, all);
				}
			}
		}
	}
	
	@Override
	public void autoFill(XRange dstRange, int fillType) {
		synchronized (_sheet) {
			if (_refs != null && !_refs.isEmpty() && !((XRangeImpl)dstRange).getRefs().isEmpty()) {
				//destination range allow only one contiguous reference
				if (((XRangeImpl)dstRange).getRefs().size() > 1) {
					throw new UiException("Command cannot be used on multiple selections");
				}
				final Ref srcRef = _refs.iterator().next();
				final Ref dstRef = ((XRangeImpl)dstRange).getRefs().iterator().next();
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
				final XSheet sheet = BookHelper.getSheet(_sheet, refSheet);
				clearContents(sheet, tRow, lCol, bRow, rCol);
			}
		}
	}
	
	private void clearContents(XSheet sheet, int tRow, int lCol, int bRow, int rCol) {
		final Set<Ref> last = new HashSet<Ref>();
		final Set<Ref> all = new HashSet<Ref>();
		
		//in clear case, should just look the existed cell
		int firstRow = Math.max(tRow,sheet.getFirstRowNum());
		int lastRow = Math.min(bRow,sheet.getLastRowNum());
		for(int r = firstRow; r <= lastRow; ++r) {
			Row row = sheet.getRow(r);
			if(row==null)
				continue;
			int firstCol = Math.max(lCol,row.getFirstCellNum());
			int lastCol = Math.min(rCol, row.getLastCellNum());
			for (int c = firstCol; c <= lastCol; ++c) {
				final Cell cell = row.getCell(c);
				if(cell==null)
					continue;
				final Set<Ref>[] refs = BookHelper.clearCell(cell);
				if (refs != null) {
					last.addAll(refs[0]);
					all.addAll(refs[1]);
				}
			}
		}
		final XBook book = (XBook) sheet.getWorkbook();
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
				fillRef(srcRef, dstRef, XRange.FILL_COPY);
			}
		}
	}
	
	@Override
	public void fillLeft() {
		synchronized (_sheet) {
			if (_refs != null && !_refs.isEmpty()) {
				final Ref dstRef = _refs.iterator().next();
				final Ref srcRef = new AreaRefImpl(dstRef.getTopRow(), dstRef.getRightCol(), dstRef.getBottomRow(), dstRef.getRightCol(), dstRef.getOwnerSheet());
				fillRef(srcRef, dstRef, XRange.FILL_COPY);
			}
		}
	}

	@Override
	public void fillRight() {
		synchronized (_sheet) {
			if (_refs != null && !_refs.isEmpty()) {
				final Ref dstRef = _refs.iterator().next();
				final Ref srcRef = new AreaRefImpl(dstRef.getTopRow(), dstRef.getLeftCol(), dstRef.getBottomRow(), dstRef.getLeftCol(), dstRef.getOwnerSheet());
				fillRef(srcRef, dstRef, XRange.FILL_COPY);
			}
		}
	}

	@Override
	public void fillUp() {
		synchronized (_sheet) {
			if (_refs != null && !_refs.isEmpty()) {
				final Ref dstRef = _refs.iterator().next();
				final Ref srcRef = new AreaRefImpl(dstRef.getBottomRow(), dstRef.getLeftCol(), dstRef.getBottomRow(), dstRef.getRightCol(), dstRef.getOwnerSheet());
				fillRef(srcRef, dstRef, XRange.FILL_COPY);
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
						final XSheet sheet = BookHelper.getSheet(_sheet, refSheet);
						final int tRow = ref.getTopRow();
						final int bRow = ref.getBottomRow();
						final Set<Ref> refs = BookHelper.setRowsHidden(sheet, tRow, bRow, hidden);
						all.addAll(refs);
					} else if (ref.isWholeColumn()) {
						final RefSheet refSheet = ref.getOwnerSheet();
						final XSheet sheet = BookHelper.getSheet(_sheet, refSheet);
						final int lCol = ref.getLeftCol();
						final int rCol = ref.getRightCol();
						final Set<Ref> refs = BookHelper.setColumnsHidden(sheet, lCol, rCol, hidden);
						all.addAll(refs);
					}
				}
				if (!all.isEmpty()) {
					final XBook book = (XBook) _sheet.getWorkbook();
					BookHelper.notifySizeChanges(book, all);
				}
			}
		}
	}
	
	@Override
	public void setDisplayGridlines(boolean show) {
		synchronized (_sheet) {
			final Ref ref = getRefs().iterator().next();
			final XSheet sheet = BookHelper.getSheet(_sheet, ref.getOwnerSheet());
			final Set<Ref> all = new HashSet<Ref>(); 
			final boolean old = sheet.isDisplayGridlines();
			if (old != show) {
				sheet.setDisplayGridlines(show);
				//sheet is important, row, column is not in this event
				all.add(ref); 
			}
			if (!all.isEmpty()) {
				final XBook book = (XBook) sheet.getWorkbook();
				BookHelper.notifyGridlines(book, all, show);
			}
		}
	}
	
	@Override
	public void protectSheet(String password) {
		synchronized (_sheet) {
			final Ref ref = getRefs().iterator().next();
			final XSheet sheet = BookHelper.getSheet(_sheet, ref.getOwnerSheet());
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
				final XBook book = (XBook) sheet.getWorkbook();
				BookHelper.notifyProtectSheet(book, all, password);
			}
		}
	}

	@Override
	public XAreas getAreas() {
		final AreasImpl areas = new AreasImpl();
		if (getRefs().size() == 1) {
			areas.addArea(this);
		} else {
			for(Ref ref : getRefs()) {
				final XRange rng = refToRange(ref);
				areas.addArea(rng);
			}
		}
		return areas;
	}
	
	private XRange refToRange(Ref ref) {
		return new XRangeImpl(ref, _sheet);
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
	public XRange getColumns() {
		final Ref ref = getRefs().iterator().next();
		final XSheet sheet = BookHelper.getSheet(_sheet, ref.getOwnerSheet());
		final int col1 = ref.getLeftCol();
		final int col2 = ref.getRightCol();
		final XBook book = (XBook) sheet.getWorkbook();
		return new XRangeImpl(0, col1, book.getSpreadsheetVersion().getLastRowIndex(), col2, sheet, sheet);
	}
	
	@Override
	public XRange getRows() {
		final Ref ref = getRefs().iterator().next();
		final XSheet sheet = BookHelper.getSheet(_sheet, ref.getOwnerSheet());
		final int row1 = ref.getTopRow();
		final int row2 = ref.getBottomRow();
		final XBook book = (XBook) sheet.getWorkbook();
		return new XRangeImpl(row1, 0, row2, book.getSpreadsheetVersion().getLastColumnIndex(), sheet, sheet);
	}
	
	@Override
	public XRange getDependents() {
		synchronized (_sheet) {
			final Ref ref = getRefs().iterator().next();
			final XSheet sheet = BookHelper.getSheet(_sheet, ref.getOwnerSheet());
			final int row = ref.getTopRow();
			final int col = ref.getLeftCol();
			final RefSheet refSheet = ref.getOwnerSheet();
			Set<Ref> refs = ((RefSheetImpl)refSheet).getAllDependents(row, col);
			return refs != null && !refs.isEmpty() ?
					new XRangeImpl(refs, sheet) : XRanges.EMPTY_RANGE;
		}
	}
	
	@Override
	public XRange getDirectDependents() {
		synchronized (_sheet) {
			final Ref ref = getRefs().iterator().next();
			final XSheet sheet = BookHelper.getSheet(_sheet, ref.getOwnerSheet());
			final int row = ref.getTopRow();
			final int col = ref.getLeftCol();
			final RefSheet refSheet = ref.getOwnerSheet();
			Set<Ref> refs = ((RefSheetImpl)refSheet).getDirectDependents(row, col);
			return refs != null && !refs.isEmpty() ?
					new XRangeImpl(refs, sheet) : XRanges.EMPTY_RANGE;
		}
	}
	
	@Override
	public XRange getPrecedents() {
		synchronized (_sheet) {
			final Ref ref = getRefs().iterator().next();
			final XSheet sheet = BookHelper.getSheet(_sheet, ref.getOwnerSheet());
			final int row = ref.getTopRow();
			final int col = ref.getLeftCol();
			final RefSheet refSheet = ref.getOwnerSheet();
			Set<Ref> refs = refSheet.getAllPrecedents(row, col);
			return refs != null && !refs.isEmpty() ?
					new XRangeImpl(refs, sheet) : XRanges.EMPTY_RANGE;
		}
		
	}
	
	@Override
	public XRange getDirectPrecedents() {
		synchronized (_sheet) {
			final Ref ref = getRefs().iterator().next();
			final XSheet sheet = BookHelper.getSheet(_sheet, ref.getOwnerSheet());
			final int row = ref.getTopRow();
			final int col = ref.getLeftCol();
			final RefSheet refSheet = ref.getOwnerSheet();
			Set<Ref> refs = refSheet.getDirectPrecedents(row, col);
			return refs != null && !refs.isEmpty() ?
					new XRangeImpl(refs, sheet) : XRanges.EMPTY_RANGE;
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
			final XBook book = (XBook)cell.getSheet().getWorkbook();
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
	public XRange getOffset(int rowOffset, int colOffset) {
		if (rowOffset == 0 && colOffset == 0) { //no offset, return this
			return this;
		}
		if (_refs != null && !_refs.isEmpty()) {
			final SpreadsheetVersion ver = ((XBook)_sheet.getWorkbook()).getSpreadsheetVersion();
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
				return XRanges.EMPTY_RANGE;
			} else{
				return new XRangeImpl(nrefs, _sheet);
			}
		}
		return XRanges.EMPTY_RANGE;
	}

	//returns the largest square range of this sheet that contains non-blank cells
	private CellRangeAddress getLargestRange(XSheet sheet) {
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
	public XRange getCurrentRegion() {
		synchronized (_sheet) {
			final Ref ref = getRefs().iterator().next();
			final int row = ref.getTopRow();
			final int col = ref.getLeftCol();
			CellRangeAddress cra = getCurrentRegion(_sheet, row, col);
			return cra == null ? 
					new XRangeImpl(row, col, _sheet, _sheet) :
					new XRangeImpl(cra.getFirstRow(), cra.getFirstColumn(), cra.getLastRow(), cra.getLastColumn(), _sheet, _sheet);
		}
	}
	
	private int[] getCellMinMax(XSheet sheet, int row, int col) {
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
	private int[] getRowMinMax(XSheet sheet, Row rowobj, int minc, int maxc) {
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
	
	//given row return the maximum range
	private CellRangeAddress getRowCurrentRegion(XSheet sheet, int topRow, int btmRow) {
		int minc = 0;
		int maxc = 0;
		int minr = topRow;
		int maxr = btmRow;
		final Row roworg = sheet.getRow(topRow);
		for (int c = minc; c <= roworg.getLastCellNum(); c++) {
			boolean foundMax = false;
			for (int r = minr + 1; r <= sheet.getLastRowNum(); r++) {
				int[] cellMinMax = getCellMinMax(sheet, r, c);
				if (cellMinMax == null && r >= btmRow) {
					break;
				}
				if (cellMinMax != null) {
					foundMax = true;
					maxr = Math.max(maxr, cellMinMax[3]);
				}
			}
			if (foundMax) {
				maxc = c;
			}
		}

		return new CellRangeAddress(minr, maxr, minc, maxc);
	}
	
	/**
	 * Search adjacent cells and return a range with non-blank cells based on a given cell.
	 * It searches row by row to find minimal and maximal non-blank columns. 
	 * @param sheet target searching sheet
	 * @param row starting cell's row index
	 * @param col starting cell's column index
	 * @return
	 */
	private CellRangeAddress getCurrentRegion(XSheet sheet, int row, int col) {
		int minNonBlankColumn = col;
		int maxNonBlankColumn = col;
		int minNonBlankRow = Integer.MAX_VALUE;
		int maxNonBlankRow = -1;
		final Row roworg = sheet.getRow(row);
		final int[] leftTopRightBottom = getRowMinMax(sheet, roworg, minNonBlankColumn, maxNonBlankColumn);
		if (leftTopRightBottom != null) {
			minNonBlankColumn = leftTopRightBottom[0];
			minNonBlankRow = leftTopRightBottom[1];
			maxNonBlankColumn = leftTopRightBottom[2];
			maxNonBlankRow = leftTopRightBottom[3];
		}
		
		int rowUp = row > 0 ? row - 1 : row;
		int rowDown = row + 1;
		
		boolean stopFindingUp = rowUp == row;
		boolean stopFindingDown = false;
		do {
			//for row above
			if (!stopFindingUp) {
				final Row rowu = sheet.getRow(rowUp);
				final int[] upperRowLeftTopRightBottom = getRowMinMax(sheet, rowu, minNonBlankColumn, maxNonBlankColumn);
				if (upperRowLeftTopRightBottom != null) {
					if (minNonBlankColumn != upperRowLeftTopRightBottom[0] || maxNonBlankColumn != upperRowLeftTopRightBottom[2]) {  //minc or maxc changed!
						stopFindingDown = false;
						minNonBlankColumn = upperRowLeftTopRightBottom[0];
						maxNonBlankColumn = upperRowLeftTopRightBottom[2];
					}
					if (minNonBlankRow > upperRowLeftTopRightBottom[1]) {
						minNonBlankRow = upperRowLeftTopRightBottom[1];
					}
					if (rowUp > 0) {
						--rowUp;
					} else {
						stopFindingUp = true; //no more row above!
					}
				} else { //blank row
					stopFindingUp = true;
				}
			}

			//for row below
			if (!stopFindingDown) {
				final Row rowd = sheet.getRow(rowDown);
				final int[] downRowLeftTopRightBottom = getRowMinMax(sheet, rowd, minNonBlankColumn, maxNonBlankColumn);
				if (downRowLeftTopRightBottom != null) {
					if (minNonBlankColumn != downRowLeftTopRightBottom[0] || maxNonBlankColumn != downRowLeftTopRightBottom[2]) { //minc and maxc changed
						stopFindingUp = false;
						minNonBlankColumn = downRowLeftTopRightBottom[0];
						maxNonBlankColumn = downRowLeftTopRightBottom[2];
					}
					if (maxNonBlankRow < downRowLeftTopRightBottom[3]) {
						maxNonBlankRow = downRowLeftTopRightBottom[3];
					}
					++rowDown;
				} else { //blank row
					stopFindingDown = true;
				}
			}
		} while(!stopFindingUp || !stopFindingDown);
		
		if (minNonBlankRow == Integer.MAX_VALUE && maxNonBlankRow < 0) { //all blanks in 9 cells!
			return null;
		}
		minNonBlankRow = (minNonBlankRow == Integer.MAX_VALUE)? row: minNonBlankRow;
		maxNonBlankRow = (maxNonBlankRow == -1)? row: maxNonBlankRow;
		return new CellRangeAddress(minNonBlankRow, maxNonBlankRow, minNonBlankColumn, maxNonBlankColumn);
	}

	// ZSS-246: give an API for user checking the auto-filtering range before applying it.
	public XRange findAutoFilterRange() {
		
		//The logic to decide the actual affected range to implement autofilter:
		//If it's a multiple cell range, it's the range intersect with largest range of the sheet.
		//If it's a single cell range, it has to be extend to a continuous range by looking up the near 8 cells of the single cell.
		CellRangeAddress currentArea = new CellRangeAddress(getRow(), getLastRow(), getColumn(), getLastColumn());
		final Ref ref = getRefs().iterator().next();
		
		//ZSS-199
		if (ref.isWholeRow()) {
			//extend to a continuous range from the top row
			CellRangeAddress cra = getRowCurrentRegion(_sheet, ref.getTopRow(), ref.getBottomRow());
			return cra != null ? XRanges.range(_sheet, cra.getFirstRow(), cra.getFirstColumn(), cra.getLastRow(), cra.getLastColumn()) : null; 
			
		} else if (BookHelper.isOneCell(_sheet, currentArea)) {
			//only one cell selected(include merged one), try to look the max range surround by blank cells 
			CellRangeAddress cra = getCurrentRegion(_sheet, getRow(), getColumn());
			return cra != null ? XRanges.range(_sheet, cra.getFirstRow(), cra.getFirstColumn(), cra.getLastRow(), cra.getLastColumn()) : null; 
			
		} else {
			CellRangeAddress largeRange = getLargestRange(_sheet); //get the largest range that contains non-blank cells
			if (largeRange == null) {
				return null;
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
				return null;
			}
			return XRanges.range(_sheet, top, left, bottom, right); 
		}
	}
	
	//TODO:
	private static final String ALL_BLANK_MSG = "Cannot find the range. Please select a cell within the range and try again!";
	@Override
	public AutoFilter autoFilter() {
		//we didn't support autofilter in 2003 since 3.0.0
		//and it cause ZSS-408 Cannot save 2003 format if the file contains auto filter configuration.
		if(_sheet instanceof HSSFSheet){
			throw new UnsupportedOperationException("filter in HSSF(2003) is not supported yet");
		}
		synchronized (_sheet) {
			CellRangeAddress affectedArea;
			final Ref ref = getRefs().iterator().next();
			if(_sheet.isAutoFilterMode()){
				affectedArea = _sheet.removeAutoFilter();
				if (affectedArea !=null){
					XRange unhideArea = XRanges.range(_sheet,
							affectedArea.getFirstRow(),affectedArea.getFirstColumn(),
							affectedArea.getLastRow(),affectedArea.getLastColumn());
					unhideArea.getRows().setHidden(false);
					BookHelper.notifyAutoFilterChange(ref,false);
				}
			} else {
				XRange r = findAutoFilterRange(); // ZSS-246: move the original code to a new API for checking
				if(r != null) {
					affectedArea = new CellRangeAddress(r.getRow(), r.getLastRow(), r.getColumn(), r.getLastColumn());
					_sheet.setAutoFilter(affectedArea);
					BookHelper.notifyAutoFilterChange(ref,true);
				} else {
					throw new RuntimeException(ALL_BLANK_MSG);
				}
			}
			 
			if (affectedArea !=null){
				//I have to know the top row area to show/remove the combo button
				//changed by this autofilter action, it's not the same area
				//and then send ON_CONTENTS_CHANGE event
				XRangeImpl buttonChange = (XRangeImpl) XRanges.range(_sheet, 
						affectedArea.getFirstRow(),affectedArea.getFirstColumn(),
						affectedArea.getFirstRow(),affectedArea.getLastColumn());

				BookHelper.notifyBtnChanges(new HashSet<Ref>(buttonChange.getRefs()));
			}
			
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
					final XRangeImpl rng = (XRangeImpl) new XRangeImpl(r, left, r, right, _sheet, _sheet); 
					all.addAll(rng.getRefs());
					rng.getRows().setHidden(false); //unhide
				}
			}
	
			BookHelper.notifyCellChanges(_sheet.getBook(), all); //unhidden row must reevaluate
			
			//update button
			final XRangeImpl buttonChange = (XRangeImpl) XRanges.range(_sheet, row1, col1, row1, col2);
			BookHelper.notifyBtnChanges(new HashSet<Ref>(buttonChange.getRefs()));
		}
	}
	
	/**
	 * Reapply filters with their existing criteria. It also re-define filtering
	 * range based on original one. If all blanks cells are found
	 * after finding, it just disables auto filter (Reference Excel's behavior). 
	 */
	@Override
	public void applyFilter() {
		synchronized (_sheet) {
			//ZSS-280
			AutoFilter oldFilter = _sheet.getAutoFilter();
			
			if (!_sheet.isAutoFilterMode()) { //no criteria is applied
				return;
			}
			//copy filtering criteria
			int firstRow = oldFilter.getRangeAddress().getFirstRow(); //first row is header
			int firstColumn = oldFilter.getRangeAddress().getFirstColumn();
			//backup column data because getting from removed auto filter will cause XmlValueDisconnectedException
			List<Object[]> originalFilteringColumns = new ArrayList<Object[]>();
			if (oldFilter.getFilterColumns() != null){ //has applied some criteria
				for (FilterColumn filterColumn : oldFilter.getFilterColumns()){
					Object[] filterColumnData = new Object[4];
					filterColumnData[0] = filterColumn.getColId()+1;
					filterColumnData[1] = filterColumn.getCriteria1().toArray(new String[0]);
					filterColumnData[2] = filterColumn.getOperator();
					filterColumnData[3] = filterColumn.getCriteria2();
					originalFilteringColumns.add(filterColumnData);
				}
			}
			
			autoFilter(); //disable existing filter
			
			//re-define filtering range 
			CellRangeAddress filteringRange = getCurrentRegion(_sheet, firstRow, firstColumn);
			if (filteringRange == null){ //Don't enable auto filter if there are all blank cells
				return;
			}else{
				//enable auto filter
				_sheet.setAutoFilter(filteringRange);
				BookHelper.notifyAutoFilterChange(getRefs().iterator().next(),true);

				//apply original criteria
				for (int nCol = 0 ; nCol < originalFilteringColumns.size(); nCol ++){
					Object[] oldFilterColumn =  originalFilteringColumns.get(nCol);
					autoFilter((Integer)oldFilterColumn[0], oldFilterColumn[1], (Integer)oldFilterColumn[2], oldFilterColumn[3], null);
				}
			}
		}
	}

	/**
	 * Add a filter with criteria and apply to the column.
	 * Not supported for Excel 2003 format.
	 * @param field n-th column of all filtering columns, 1-based, e.g. 2nd column is 2 
	 * @param criteria1 string array of selected data, e.g. ["a", "d", "e"]
	 */
	@Override
	public AutoFilter autoFilter(int field, Object criteria1, int filterOp, Object criteria2, Boolean visibleDropDown) {
		//we didn't support autofilter in 2003 since 3.0.0
		//and it cause ZSS-408 Cannot save 2003 format if the file contains auto filter configuration.
		if(_sheet instanceof HSSFSheet){
			throw new UnsupportedOperationException("filter in HSSF(2003) is not supported yet");
		}		
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
						new XRangeImpl(r, col, _sheet, _sheet).getRows().setHidden(true);
					}
				} else { //candidate to be shown (other FieldColumn might still hide this row!
					final Row rowobj = _sheet.getRow(r);
					if (rowobj != null && rowobj.getZeroHeight() && canUnhide(af, fc, r, col1)) { //a hidden row and no other hidden filtering
						final int left = rowobj.getFirstCellNum();
						final int right = rowobj.getLastCellNum() - 1;
						final XRangeImpl rng = (XRangeImpl) new XRangeImpl(r, left, r, right, _sheet, _sheet); 
						all.addAll(rng.getRefs());
						rng.getRows().setHidden(false); //unhide
					}
				}
			}
			
			BookHelper.notifyCellChanges(_sheet.getBook(), all); //unhidden row must reevaluate
			
			//update button
			final XRangeImpl buttonChange = (XRangeImpl) XRanges.range(_sheet, row1, col, row1, col);
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
			final XRangeImpl rng = (XRangeImpl) XRanges.range(_sheet, anchor.getRow1(), anchor.getCol1(), anchor.getRow2(), anchor.getCol2());
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
			final XRangeImpl rng = (XRangeImpl) XRanges.range(_sheet, anchor.getRow1(), anchor.getCol1(), anchor.getRow2(), anchor.getCol2());
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
			final XRangeImpl rng = (XRangeImpl) XRanges.range(_sheet, anchor.getRow1(), anchor.getCol1(), anchor.getRow2(), anchor.getCol2());
			final Collection<Ref> refs = rng.getRefs();
			// ZSS-397: keep picture ID for notifying
			// the picture data including ID will be gone after deleting
			final String id = picture.getPictureId();
			dm.deletePicture(_sheet, picture); //must after getPreferredSize() or anchor is gone!
			if (refs != null && !refs.isEmpty()) {
				final Ref ref = refs.iterator().next();
				BookHelper.notifyPictureDelete(ref, id);
			}
		}
	}

	@Override
	public void movePicture(Picture picture, ClientAnchor anchor) {
		synchronized (_sheet) {
			DrawingManager dm = ((SheetCtrl)_sheet).getDrawingManager();
			dm.movePicture(_sheet, picture, anchor);
			final XRangeImpl rng = (XRangeImpl) XRanges.range(_sheet, anchor.getRow1(), anchor.getCol1(), anchor.getRow2(), anchor.getCol2());
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
			ZssChartX chartX = dm.getChartX(chart);
			if(chartX==null) return;
			
			dm.moveChart(_sheet, chart, anchor);
			final XRangeImpl rng = (XRangeImpl) XRanges.range(_sheet, anchor.getRow1(), anchor.getCol1(), anchor.getRow2(), anchor.getCol2());
			final Collection<Ref> refs = rng.getRefs();
			if (refs != null && !refs.isEmpty()) {
				final Ref ref = refs.iterator().next();
				BookHelper.notifyChartUpdate(ref, chartX);
			}
		}
	}

	@Override
	public void deleteChart(Chart chart) {
		synchronized (_sheet) {
			DrawingManager dm = ((SheetCtrl)_sheet).getDrawingManager();
			ClientAnchor anchor = chart.getPreferredSize();
			final XRangeImpl rng = (XRangeImpl) XRanges.range(_sheet, anchor.getRow1(), anchor.getCol1(), anchor.getRow2(), anchor.getCol2());
			final Collection<Ref> refs = rng.getRefs();
			// ZSS-358: keep chart ID for notifying; must assume that chart data was gone.
			final String id = chart.getChartId();
			dm.deleteChart(_sheet, chart); //must after getPreferredSize() or anchor is gone!
			if (refs != null && !refs.isEmpty()) {
				final Ref ref = refs.iterator().next();
				BookHelper.notifyChartDelete(ref, id);
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
	public void createSheet(String name) {
		synchronized (_sheet.getBook()) {
			Ref ref = _refs != null && !_refs.isEmpty() ? _refs.iterator().next() : null;
			if (ref != null) {
				final XBook book = _sheet.getBook();
				if (book != null) {
					if (Strings.isBlank(name)) {
						book.createSheet();
					} else {
						book.createSheet(name);
					}
					BookHelper.notifyCreateSheet(ref, name);
				}
			}
		}
	}

	@Override
	public void setSheetName(String name) {
		synchronized (_sheet.getBook()) {
			Ref ref = _refs != null && !_refs.isEmpty() ? _refs.iterator().next() : null;
			if (ref != null) {
				final XBook book = _sheet.getBook();
				if (book != null) {
					if (!Strings.isBlank(name)) {
						final int pos= book.getSheetIndex(_sheet);
						if (pos >= 0) {
							book.setSheetName(pos, name);
							BookHelper.notifyChangeSheetName(ref, name);
						}
					}
				}
			}
		}
	}
	
	@Override
	public void setSheetOrder(int pos) {
		synchronized (_sheet.getBook()) {
			Ref ref = _refs != null && !_refs.isEmpty() ? _refs.iterator().next() : null;
			if (ref != null) {
				final XBook book = _sheet.getBook();
				if (book != null) {
					final String name = _sheet.getSheetName();
					if (!Strings.isBlank(name)) {
						book.setSheetOrder(name, pos);
						book.getFormulaEvaluator().clearAllCachedResultValues(); // ZSS-492: don't forget to clear cache of evaluator
						BookHelper.notifyChangeSheetOrder(ref, name);
					}
				}
			}
		}
	}

	@Override
	public void deleteSheet() {
		synchronized (_sheet.getBook()) {
			Ref ref = _refs != null && !_refs.isEmpty() ? _refs.iterator().next() : null;
			if (ref != null) {
				final XBook book = _sheet.getBook();
				if (book != null) {
					final int index = book.getSheetIndex(_sheet);
					if (index != -1) {
						final int sheetCount = book.getNumberOfSheets();
						if (sheetCount == 1) {
							Messagebox.show("A workbook must contain at least one visible worksheet");
							//ZSS-493, return do noting or will get exception.
							return;
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
	
	public boolean isWholeRow(){
		Ref ref = _refs.iterator().next();
		return ref.isWholeRow();
	}
	
	public boolean isWholeColumn(){
		Ref ref = _refs.iterator().next();
		return ref.isWholeColumn();
	}
	
	public boolean isWholeSheet(){
		Ref ref = _refs.iterator().next();
		return ref.isWholeSheet();
	}
	
	public void setFreezePanel(int rowfreeze, int columnfreeze){
		synchronized (_sheet.getBook()) {
			Ref ref = _refs != null && !_refs.isEmpty() ? _refs.iterator().next() : null;
			if (ref != null) {
				BookHelper.setFreezePanel(_sheet, rowfreeze, columnfreeze);
				BookHelper.notifyFreezeSheet(ref, new Object[] {rowfreeze, columnfreeze});
			}
		}
	}
}
