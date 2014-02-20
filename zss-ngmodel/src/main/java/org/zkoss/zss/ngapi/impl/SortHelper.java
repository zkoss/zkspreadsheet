package org.zkoss.zss.ngapi.impl;

import java.io.Serializable;
import java.util.*;
import java.util.Map.Entry;

import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.engine.impl.*;
import org.zkoss.zss.model.sys.*;
import org.zkoss.zss.model.sys.impl.BookHelper;
import org.zkoss.zss.model.sys.impl.BookHelper.*;
import org.zkoss.zss.ngapi.NRange;

/**
 * Manipulate cells according to sorting criteria and options.
 * @author Hawk
 *
 */
//porting implementation from BookHelper.sort()
public class SortHelper extends RangeHelperBase {

	public static final int SORT_NORMAL_DEFAULT = 0;
	public static final int SORT_TEXT_AS_NUMBERS = 1;
	public static final int SORT_HEADER_NO  = 0;
	public static final int SORT_HEADER_YES = 1;	
	
	public SortHelper(NRange range) {
		super(range);
	}
	
	@SuppressWarnings("unchecked")
	public static ChangeInfo sort(XSheet sheet, int tRow, int lCol, int bRow, int rCol, 
			Ref key1, boolean desc1, Ref key2, int type, boolean desc2, Ref key3, boolean desc3, int header, int orderCustom,
			boolean matchCase, boolean sortByRows, int sortMethod, int dataOption1, int dataOption2, int dataOption3) {
		//TODO type not yet imiplmented(Sort label/Sort value, for PivotTable)
		//TODO orderCustom is not implemented yet
		if (header == BookHelper.SORT_HEADER_YES) {
			if (sortByRows) {
				++lCol;
			} else {
				++tRow;
			}
		}
		if (tRow > bRow || lCol > rCol) { //nothing to sort!
			return null;
		}
		int keyCount = 0;
		if (key1 != null) {
			++keyCount;
			if (key2 != null) {
				++keyCount;
				if (key3 != null) {
					++keyCount;
				}
			}
		}
		if (keyCount == 0) {
			throw new UiException("Must specify at least the key1");
		}
		final int[] dataOptions = new int[keyCount];
		final boolean[] descs = new boolean[keyCount];
		final int[] keyIndexes = new int[keyCount];
		keyIndexes[0] = rangeToIndex(key1, sortByRows);
		descs[0] = desc1;
		dataOptions[0] = dataOption1;
		if (keyCount > 1) {
			keyIndexes[1] = rangeToIndex(key2, sortByRows);
			descs[1] = desc2;
			dataOptions[1] = dataOption2;
		}
		if (keyCount > 2) {
			keyIndexes[2] = rangeToIndex(key3, sortByRows);
			descs[2] = desc3;
			dataOptions[2] = dataOption3;
		}
		validateKeyIndexes(keyIndexes, tRow, lCol, bRow, rCol, sortByRows);
		
		final List<SortKey> sortKeys = new ArrayList<SortKey>(sortByRows ? rCol - lCol + 1 : bRow - tRow + 1);
		final int begRow = Math.max(tRow, sheet.getFirstRowNum());
		final int endRow = Math.min(bRow, sheet.getLastRowNum());
		if (sortByRows) {
			int begCol = ((XBook)sheet.getWorkbook()).getSpreadsheetVersion().getLastColumnIndex();
			int endCol = 0;
			//locate begCol/endCol of the sheet
			for (int rowNum = begRow; rowNum <= endRow; ++rowNum) {
				final Row row = sheet.getRow(rowNum);
				if (row != null) {
					begCol = Math.min(begCol, row.getFirstCellNum());
					endCol = Math.max(begCol, row.getLastCellNum() - 1);
				}
			}
			begCol = Math.max(lCol, begCol);
			endCol = Math.min(rCol, endCol);
			for (int colnum = begCol; colnum <= endCol; ++colnum) {
				final Object[] values = new Object[keyCount];
				for(int j = 0; j < keyCount; ++j) {
					final Row row = sheet.getRow(keyIndexes[j]);
					final Cell cell = row != null ? row.getCell(colnum, Row.RETURN_BLANK_AS_NULL) : null;
					final Object val = getCellObject(cell, dataOptions[j]);
					values[j] = val;
				}
				final SortKey sortKey = new SortKey(colnum, values);
				sortKeys.add(sortKey);
			}
			if (!sortKeys.isEmpty()) {
				final Comparator<SortKey> keyComparator = new KeyComparator(descs, matchCase, sortMethod, type);
				Collections.sort(sortKeys, keyComparator);
				return BookHelper.assignColumns(sheet, sortKeys, begRow, lCol, endRow, rCol);
			} else {
				return null;
			}
		} else { //sortByColumn, default case
			for (int rownum = begRow; rownum <= endRow; ++rownum) {
				final Row row = sheet.getRow(rownum);
				if (row == null) {
					continue; //nothing to sort
				}
				final Object[] values = new Object[keyCount];
				for(int j = 0; j < keyCount; ++j) {
					final Cell cell = row.getCell(keyIndexes[j], Row.RETURN_BLANK_AS_NULL);
					final Object val = getCellObject(cell, dataOptions[j]);
					values[j] = val;
				}
				final SortKey sortKey = new SortKey(rownum, values);
				sortKeys.add(sortKey);
			}
			if (!sortKeys.isEmpty()) {
				final Comparator<SortKey> keyComparator = new KeyComparator(descs, matchCase, sortMethod, type);
				Collections.sort(sortKeys, keyComparator);
				return BookHelper.assignRows(sheet, sortKeys, tRow, lCol, bRow, rCol);
			} else {
				return null;
			}
		}
	}
	
	private static Object getCellObject(Cell cell, int dataOption) {
		Object val = cell != null ? BookHelper.getCellObject(cell) : null;
		if (val instanceof RichTextString && dataOption == BookHelper.SORT_TEXT_AS_NUMBERS) {
			try {
				val = new Double((String)((RichTextString)val).getString());
			} catch(NumberFormatException ex) {
				val = new Double(0);//ignore
			}
		}
		return val;
	}
	@SuppressWarnings("unchecked")
	private static ChangeInfo  assignColumns(XSheet sheet, List<SortKey> sortKeys, int tRow, int lCol, int bRow, int rCol) {
		final int cellCount = bRow - tRow + 1;
		final Map<Integer, List<Cell>> newCols = new HashMap<Integer, List<Cell>>();  
		final Set<Ref> toEval = new HashSet<Ref>();
		final Set<Ref> affected = new HashSet<Ref>();
		final List<MergeChange> mergeChanges = new ArrayList<MergeChange>();
		final ChangeInfo changeInfo = new ChangeInfo(toEval, affected, mergeChanges);
		int j = 0;
		for(final Iterator<SortKey> it = sortKeys.iterator(); it.hasNext();++j) {
			final SortKey sortKey = it.next();
			final int oldColNum = sortKey.getIndex();
			final int newColNum = lCol + j;
			it.remove();
			if (oldColNum == newColNum) { //no move needed, skip it
				continue;
			}
			//remove cells from the old column of the Range
			final List<Cell> cells = new ArrayList<Cell>(cellCount);
			for(int k = tRow; k <= bRow; ++k) {
				final Cell cell = BookHelper.getCell(sheet, k, oldColNum);
				if (cell != null) {
					cells.add(cell);
					final Set<Ref>[] refs = BookHelper.removeCell(cell, false);
					toEval.addAll(refs[0]);
					affected.addAll(refs[1]);
				}
			}
			if (!cells.isEmpty()) {
				newCols.put(Integer.valueOf(newColNum), cells);
			}
		}
		
		//move cells
		for(Entry<Integer, List<Cell>> entry : newCols.entrySet()) {
			final int colNum = entry.getKey().intValue();
			final List<Cell> cells = entry.getValue();
			for(Cell cell : cells) {
				final int rowNum = cell.getRowIndex();
				final ChangeInfo changeInfo0 = BookHelper.copyCell(cell, sheet, rowNum, colNum, XRange.PASTE_ALL, XRange.PASTEOP_NONE, false);
				assignChangeInfo(toEval, affected, mergeChanges, changeInfo0);
			}
		}
		return changeInfo;
	}

	@SuppressWarnings("unchecked")
	private static ChangeInfo assignRows(XSheet sheet, List<SortKey> sortKeys, int tRow, int lCol, int bRow, int rCol) {
		final int cellCount = rCol - lCol + 1;
		final Map<Integer, List<Cell>> newRows = new HashMap<Integer, List<Cell>>();  
		final Set<Ref> toEval = new HashSet<Ref>();
		final Set<Ref> affected = new HashSet<Ref>();
		final List<MergeChange> mergeChanges = new ArrayList<MergeChange>();
		final ChangeInfo changeInfo = new ChangeInfo(toEval, affected, mergeChanges);
		int j = 0;
		for(final Iterator<SortKey> it = sortKeys.iterator(); it.hasNext();++j) {
			final SortKey sortKey = it.next();
			final int oldRowNum = sortKey.getIndex();
			final Row row = sheet.getRow(oldRowNum); 
			final int newRowNum = tRow + j;
			it.remove();
			if (oldRowNum == newRowNum) { //no move needed, skip it
				continue;
			}
			//remove cells from the old row of the Range
			final List<Cell> cells = new ArrayList<Cell>(cellCount);
			final int begCol = Math.max(lCol, row.getFirstCellNum());
			final int endCol = Math.min(rCol, row.getLastCellNum() - 1);
			for(int k = begCol; k <= endCol; ++k) {
				final Cell cell = row.getCell(k);
				if (cell != null) {
					cells.add(cell);
					final Set<Ref>[] refs = BookHelper.removeCell(cell, false);
					assignRefs(toEval, affected, refs);
				}
			}
			if (!cells.isEmpty()) {
				newRows.put(Integer.valueOf(newRowNum), cells);
			}
		}
		
		//move cells
		for(Entry<Integer, List<Cell>> entry : newRows.entrySet()) {
			final int rowNum = entry.getKey().intValue();
			final List<Cell> cells = entry.getValue();
			for(Cell cell : cells) {
				final int colNum = cell.getColumnIndex();
				final ChangeInfo changeInfo0 = BookHelper.copyCell(cell, sheet, rowNum, colNum, XRange.PASTE_ALL, XRange.PASTEOP_NONE, false);
				assignChangeInfo(toEval, affected, mergeChanges, changeInfo0);
			}
		}
		return changeInfo;
	}
	private static Object getCellObject(Cell cell) {
		if (cell == null) {
			return "";
		}
		int cellType = cell.getCellType();
		if (cellType == Cell.CELL_TYPE_FORMULA) {
			final XBook book = (XBook)cell.getSheet().getWorkbook();
			final CellValue cv = BookHelper.evaluate(book, cell);
			return BookHelper.getValueByCellValue(cv);
		} else {
			return BookHelper.getCellValue(cell);
		}
	}
	private static int rangeToIndex(Ref range, boolean sortByRows) {
		return sortByRows ? range.getTopRow() : range.getLeftCol();
	}
	private static void validateKeyIndexes(int[] keyIndexes, int tRow, int lCol, int bRow, int rCol, boolean sortByRows) {
		if (!sortByRows) {
			for(int j = keyIndexes.length - 1; j >= 0; --j) {
				final int keyIndex = keyIndexes[j]; 
				if (keyIndex < lCol || keyIndex > rCol) {
					throw new UiException("The given key is out of the sorting range: "+keyIndex);
				}
			}
		} else {
			for(int j = keyIndexes.length - 1; j >= 0; --j) {
				final int keyIndex = keyIndexes[j]; 
				if (keyIndex < tRow || keyIndex > bRow) {
					throw new UiException("The given key is out of the sorting range: "+keyIndex);
				}
			}
		}
	}
	
	public static class SortKey {
		final private int _index; //original row/column index
		final private Object[] _values;
		public SortKey(int index, Object[] values) {
			this._index = index;
			this._values = values;
		}
		public int getIndex() {
			return _index;
		}
		public Object[] getValues() {
			return _values;
		}
	}
	
	private static class KeyComparator implements Comparator<SortKey>, Serializable {
		final private boolean[] _descs;
		final private boolean _matchCase;
		final private int _sortMethod; //TODO byNumberOfStroks, byPinyYin
		final private int _type; //TODO PivotTable only: byLabel, byValue
		
		public KeyComparator(boolean[] descs, boolean matchCase, int sortMethod, int type) {
			_descs = descs;
			_matchCase = matchCase;
			_sortMethod = sortMethod;
			_type = type;
		}
		@Override
		public int compare(SortKey o1, SortKey o2) {
			final Object[] values1 = o1.getValues();
			final Object[] values2 = o2.getValues();
			return compare(values1, values2);
		}

		private int compare(Object[] values1, Object[] values2) {
			final int len = values1.length;
			for(int j = 0; j < len; ++j) {
				int p = compareValue(values1[j], values2[j], _descs[j]);
				if (p != 0) {
					return p;
				}
			}
			return 0;
		}
		//1. null is always sorted at the end
		//2. Error(Byte) > Boolean > String > Double
		private int compareValue(Object val1, Object val2, boolean desc) {
			if (val1 == val2) {
				return 0;
			}
			final int order1 = val1 instanceof Byte ? 4 : val1 instanceof Boolean ? 3 : val1 instanceof RichTextString ? 2 : val1 instanceof Number ? 1 : desc ? 0 : 5;
			final int order2 = val2 instanceof Byte ? 4 : val2 instanceof Boolean ? 3 : val2 instanceof RichTextString ? 2 : val2 instanceof Number ? 1 : desc ? 0 : 5;
			int ret = 0;
			if (order1 != order2) {
				ret = order1 - order2;
			} else { //order1 == order2
				switch(order1) {
				case 4: //error, no order among different errors
					ret = 0;
					break;
				case 3: //Boolean
					ret = ((Boolean)val1).compareTo((Boolean)val2);
					break;
				case 2: //RichTextString
					ret = compareString(((RichTextString)val1).getString(), ((RichTextString)val2).getString());
					break;
				case 1: //Double
					ret =((Double)val1).compareTo((Double)val2);
					break;
				default:
					throw new UiException("Unknown value type: "+val1.getClass());
				}
			}
			return desc ? -ret : ret;
		}
		private int compareString(String s1, String s2) {
			return _matchCase ? compareString0(s1, s2) : s1.compareToIgnoreCase(s2);
		}
		//bug 59 Sort with case sensitive should be in special spreadsheet order
		private int compareString0(String s1, String s2) {
			final int len1 = s1.length();
			final int len2 = s2.length();
			final int len = len1 > len2 ? len2 : len1;
			for (int j = 0; j < len; ++j) {
				final int ret = compareChar(s1.charAt(j), s2.charAt(j));
				if ( ret != 0) {
					return ret;
				}
			}
			return len1 - len2;
		}
		private int compareChar(char ch1, char ch2) {
			final char uch1 = Character.toUpperCase(ch1);
			final char uch2 = Character.toUpperCase(ch2);
			return uch1 == uch2 ? 
					(ch2 - ch1) : //yes, a < A
					(uch1 - uch2); //yes, a < b, a < B, A < b, and A < B
		}
	}	
}
