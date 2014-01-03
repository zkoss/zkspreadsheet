package org.zkoss.zss.ngapi.impl;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.XRanges;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.model.sys.impl.BookHelper;
import org.zkoss.zss.model.sys.impl.XRangeImpl;
import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngmodel.NSheet;

/**
 * to help seach the max data region
 * @author dennis
 *
 */
//these code if from XRangeImpl,  
public class DataRegionHelper {

	NSheet sheet;
	NRange range;
	
	public DataRegionHelper(NSheet sheet, NRange range){
		this.sheet = sheet;
		this.range = range;
	}

	public int getRow() {
		return range.getRow();
	}

	public int getColumn() {
		return range.getColumn();
	}

	public int getLastRow() {
		return range.getLastRow();
	}

	public int getLastColumn() {
		return range.getLastColumn();
	}

	public boolean isWholeRow(){
		return range.isWholeRow();
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
	
	public static boolean isOneCell(XSheet sheet, CellRangeAddress rng) {
		if (rng.getNumberOfCells() == 1) {
			return true;
		}
		final int l = rng.getFirstColumn();
		final int t = rng.getFirstRow();
		final int r = rng.getLastColumn();
		final int b = rng.getLastRow();
		
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
			final CellRangeAddress ref = sheet.getMergedRegion(i);
			final int top = ref.getFirstRow();
	        final int bottom = ref.getLastRow();
	        final int left = ref.getFirstColumn();
	        final int right = ref.getLastColumn();
	    
	        if (l == left && t == top && r == right && b == bottom) {
	        	return true;
	        }
		}
		return false;
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
	
}
