package org.zkoss.zss.ngapi.impl;

import java.io.Serializable;
import java.util.*;
import java.util.Map.Entry;

import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngapi.impl.imexp.BookHelper;
import org.zkoss.zss.ngmodel.*;
import org.zkoss.zss.ngmodel.PasteOption.PasteType;

/**
 * Manipulate cells according to sorting criteria and options.
 * @author Hawk
 *
 */
//porting implementation from BookHelper.sort()
public class SortHelper extends RangeHelperBase {

	private static final PasteOption PASTE_ALL_OPTION = new PasteOption();
	
	public SortHelper(NRange range) {
		super(range);
		PASTE_ALL_OPTION.setPasteType(PasteType.ALL);
	}
	
	/**
	 * 
	 * @param sheet
	 * @param tRow selection range to sort
	 * @param lCol selection range to sort
	 * @param bRow selection range to sort
	 * @param rCol selection range to sort
	 * @param key1
	 * @param desc1
	 * @param key2
	 * @param type
	 * @param desc2
	 * @param key3
	 * @param desc3
	 * @param header
	 * @param orderCustom
	 * @param matchCase
	 * @param sortByRows
	 * @param sortMethod
	 * @param dataOption1 BookHelper.SORT_TEXT_AS_NUMBERS, BookHelper.SORT_NORMAL_DEFAULT
	 * @param dataOption2
	 * @param dataOption3
	 */
	public void sort(NSheet sheet, int tRow, int lCol, int bRow, int rCol, 
			NRange key1, boolean desc1, NRange key2, int type, boolean desc2, NRange key3, boolean desc3, int header, int orderCustom,
			boolean matchCase, boolean sortByRows, int sortMethod, int dataOption1, int dataOption2, int dataOption3) {
		//TODO type not yet implemented(Sort label/Sort value, for PivotTable)
		//TODO orderCustom is not implemented yet
		if (header == BookHelper.SORT_HEADER_YES) {
			if (sortByRows) {
				++lCol;
			} else {
				++tRow;
			}
		}
		if (tRow > bRow || lCol > rCol) { //nothing to sort!
			return ;
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
//		final int begRow = Math.max(tRow, sheet.getFirstRowNum());
//		final int endRow = Math.min(bRow, sheet.getLastRowNum());
		final int begRow = Math.max(tRow, sheet.getStartRowIndex());
		final int endRow = Math.min(bRow, sheet.getEndRowIndex());
		if (sortByRows) { //keyIndex contains row index
			int begCol = sheet.getBook().getMaxColumnIndex();
			int endCol = 0;
			//locate begCol/endCol of the sheet
			for (int rowNum = begRow; rowNum <= endRow; ++rowNum) {
				final NRow row = sheet.getRow(rowNum);
				if (row != null) {
//					begCol = Math.min(begCol, row.getFirstCellNum());
//					endCol = Math.max(begCol, row.getLastCellNum() - 1);
					begCol = Math.min(begCol, sheet.getStartCellIndex(rowNum));
					endCol = Math.max(begCol, sheet.getEndCellIndex(rowNum) - 1);
				}
			}
			begCol = Math.max(lCol, begCol);
			endCol = Math.min(rCol, endCol);
			for (int colnum = begCol; colnum <= endCol; ++colnum) {
				Object[] values = new Object[keyCount];
				for(int j = 0; j < keyCount; ++j) {
					NRow row = sheet.getRow(keyIndexes[j]);
//					final Cell cell = row != null ? row.getCell(colnum, Row.RETURN_BLANK_AS_NULL) : null;
					NCell cell = sheet.getCell(keyIndexes[j], colnum);
					Object val = getCellObject(cell, dataOptions[j]);
					values[j] = val;
				}
				SortKey sortKey = new SortKey(colnum, values);
				sortKeys.add(sortKey);
			}
			if (!sortKeys.isEmpty()) {
				final Comparator<SortKey> keyComparator = new KeyComparator(descs, matchCase, sortMethod, type);
				Collections.sort(sortKeys, keyComparator);
				 assignColumns(sheet, sortKeys, begRow, lCol, endRow, rCol);
			} else {
				return ;
			}
		} else { //sortByColumn, default case , keyIndex contains column index
			for (int rownum = begRow; rownum <= endRow; ++rownum) {
				final NRow row = sheet.getRow(rownum);
				if (row == null) {
					continue; //nothing to sort
				}
				final Object[] values = new Object[keyCount];
				for(int j = 0; j < keyCount; ++j) {
//					final Cell cell = row.getCell(keyIndexes[j], Row.RETURN_BLANK_AS_NULL);
					final NCell cell = sheet.getCell(rownum, keyIndexes[j]);
					final Object val = getCellObject(cell, dataOptions[j]);
					values[j] = val;
				}
				final SortKey sortKey = new SortKey(rownum, values);
				sortKeys.add(sortKey);
			}
			if (!sortKeys.isEmpty()) {
				final Comparator<SortKey> keyComparator = new KeyComparator(descs, matchCase, sortMethod, type);
				Collections.sort(sortKeys, keyComparator);
				assignRows(sheet, sortKeys, tRow, lCol, bRow, rCol);
			} else {
				return ;
			}
		}
	}
	
	private int rangeToIndex(NRange range, boolean sortByRows) {
		return sortByRows ? range.getRow() : range.getColumn();
	}
	
	private void validateKeyIndexes(int[] keyIndexes, int tRow, int lCol, int bRow, int rCol, boolean sortByRows) {
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
	
	@SuppressWarnings("unchecked")
	private static void  assignColumns(NSheet sheet, List<SortKey> sortKeys, int tRow, int lCol, int bRow, int rCol) {
		final int cellCount = bRow - tRow + 1;
		final Map<Integer, List<NCell>> newCols = new HashMap<Integer, List<NCell>>(); //key: new column index after sorting 
//		final Set<Ref> toEval = new HashSet<Ref>();
//		final Set<Ref> affected = new HashSet<Ref>();
//		final List<MergeChange> mergeChanges = new ArrayList<MergeChange>();
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
			final List<NCell> cells = new ArrayList<NCell>(cellCount);
			for(int k = tRow; k <= bRow; ++k) {
				NCell cell = sheet.getCell(k, oldColNum);
				if (cell != null) {
					cells.add(cell);
//					final Set<Ref>[] refs = BookHelper.removeCell(cell, false);
//					toEval.addAll(refs[0]);
//					affected.addAll(refs[1]);
				}
			}
			if (!cells.isEmpty()) {
				newCols.put(Integer.valueOf(newColNum), cells);
			}
		}
		
		//move cells
		for(Entry<Integer, List<NCell>> entry : newCols.entrySet()) {
			final int colNum = entry.getKey().intValue();
			final List<NCell> cells = entry.getValue();
			for(NCell cell : cells) {
				final int rowNum = cell.getRowIndex();
//				BookHelper.copyCell(cell, sheet, rowNum, colNum, XRange.PASTE_ALL, XRange.PASTEOP_NONE, false);
			}
		}
	}

	/**
	 * Change order of rows according to sorting result.
	 * @param sheet
	 * @param sortKeys
	 * @param tRow
	 * @param lCol
	 * @param bRow
	 * @param rCol
	 * @return
	 */
	@SuppressWarnings("unchecked")
	private static void assignRows(NSheet sheet, List<SortKey> sortKeys, int tRow, int lCol, int bRow, int rCol) {
		final int cellCount = rCol - lCol + 1;
		final Map<Integer, List<NCell>> newRows = new HashMap<Integer, List<NCell>>(); //key: new row index after sorting 
//		final Set<Ref> toEval = new HashSet<Ref>();
//		final Set<Ref> affected = new HashSet<Ref>();
		int j = 0;
		for(final Iterator<SortKey> it = sortKeys.iterator(); it.hasNext();++j) {
			final SortKey sortKey = it.next();
			final int oldRowNum = sortKey.getIndex();
			final NRow row = sheet.getRow(oldRowNum); 
			final int newRowNum = tRow + j;
			it.remove();
			if (oldRowNum == newRowNum) { //no move needed, skip it
				continue;
			}
			//remove cells from the old row of the Range
			final List<NCell> cells = new ArrayList<NCell>(cellCount);
			final int begCol = Math.max(lCol, sheet.getStartCellIndex(oldRowNum));
			final int endCol = Math.min(rCol, sheet.getEndCellIndex(oldRowNum) - 1);
			for(int k = begCol; k <= endCol; ++k) {
				final NCell cell = sheet.getCell(oldRowNum, k);
				if (cell != null) {
					cells.add(cell);
//					final Set<Ref>[] refs = BookHelper.removeCell(cell, false);
//					assignRefs(toEval, affected, refs);
				}
			}
			if (!cells.isEmpty()) {
				newRows.put(Integer.valueOf(newRowNum), cells);
			}
//			CellRegion sourceRow = new CellRegion(oldRowNum,lCol ,oldRowNum, rCol);
//			CellRegion destinationRow = new CellRegion(newRowNum,lCol ,newRowNum, rCol);
//			sheet.pasteCell(new SheetRegion(sheet, sourceRow), destinationRow, PASTE_ALL_OPTION);
		}
		
		//move cells
		for(Entry<Integer, List<NCell>> entry : newRows.entrySet()) {
			final int newRowIndex = entry.getKey().intValue();
			final List<NCell> cells = entry.getValue();
//			for(Cell cell : cells) {
//				final int colNum = cell.getColumnIndex();
//				BookHelper.copyCell(cell, sheet, newRowIndex, colNum, XRange.PASTE_ALL, XRange.PASTEOP_NONE, false);
//				assignChangeInfo(toEval, affected, mergeChanges, changeInfo0);
//			}
		}
	}
	
	//convert cell sorting data upon data option
	private Object getCellObject(NCell cell, int dataOption) {
		Object val = cell.getValue();
		if (val instanceof RichTextString && dataOption == BookHelper.SORT_TEXT_AS_NUMBERS) {
			try {
				val = new Double((String)((RichTextString)val).getString());
			} catch(NumberFormatException ex) {
				val = new Double(0);//ignore
			}
		}
		return val;
	}
	
	//TODO getCellValue()
	private Object getCellObject(NCell cell) {
//		if (cell == null) {
//			return "";
//		}
//		int cellType = cell.getCellType();
//		if (cellType == Cell.CELL_TYPE_FORMULA) {
//			final XBook book = (XBook)cell.getSheet().getWorkbook();
//			final CellValue cv = BookHelper.evaluate(book, cell);
//			return BookHelper.getValueByCellValue(cv);
//		} else {
//			return BookHelper.getCellValue(cell);
//		}
		return null;
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
