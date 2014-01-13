package org.zkoss.zss.ngapi.impl;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.RichTextString;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngapi.impl.ChangeInfo;
import org.zkoss.zss.ngapi.impl.RangeHelperBase;

public class MergeHelper extends RangeHelperBase{

	public MergeHelper(NRange range) {
		super(range);
	}
	
	public static ChangeInfo unMerge(XSheet sheet, int tRow, int lCol, int bRow, int rCol,boolean overlapped) {
		final RefSheet refSheet = BookHelper.getRefSheet((XBook)sheet.getWorkbook(), sheet);
		final List<MergeChange> changes = new ArrayList<MergeChange>(); 
		for(int j = sheet.getNumMergedRegions() - 1; j >= 0; --j) {
        	final CellRangeAddress merged = sheet.getMergedRegion(j);
        	
        	final int firstCol = merged.getFirstColumn();
        	final int lastCol = merged.getLastColumn();
        	final int firstRow = merged.getFirstRow();
        	final int lastRow = merged.getLastRow();
        	
        	// ZSS-395 unmerge when any cell overlap with merged region
        	// ZSS-412 use a flag to decide to check overlap or not.
        	if( (overlapped && overlap(firstRow, firstCol, lastRow, lastCol, tRow, lCol, bRow, rCol)) || 
        			(!overlapped && contain(tRow, lCol, bRow, rCol,firstRow, firstCol, lastRow, lastCol)) ) {
				changes.add(new MergeChange(new AreaRefImpl(firstRow, firstCol, lastRow, lastCol, refSheet), null));
				sheet.removeMergedRegion(j);
        	}
		}
		return new ChangeInfo(null, null, changes);
	}
	
	//a b are overlapped.
	private static boolean overlap(int aTopRow, int aLeftCol, int aBottomRow, int aRightCol,
			int bTopRow, int bLeftCol, int bBottomRow, int bRightCol) {
		
		boolean xOverlap = between(aLeftCol, bLeftCol, bRightCol) || between(bLeftCol, aLeftCol, aRightCol);
		boolean yOverlap = between(aTopRow, bTopRow, bBottomRow) || between(bTopRow, aTopRow, aBottomRow);
		
		return xOverlap && yOverlap;
	}
	
	//a contains b
	private static boolean contain(int aTopRow, int aLeftCol, int aBottomRow, int aRightCol,
			int bTopRow, int bLeftCol, int bBottomRow, int bRightCol){
		return aLeftCol <= bLeftCol && aRightCol >= bRightCol 
        		&& aTopRow <= bTopRow && aBottomRow >= bBottomRow;
	}
	
	private static boolean between(int value, int min, int max) {
		return (value >= min) && (value <= max);
	}
	
	/*
	 * Merge the specified range per the given tRow, lCol, bRow, rCol.
	 * 
	 * @param sheet sheet
	 * @param tRow top row
	 * @param lCol left column
	 * @param bRow bottom row
	 * @param rCol right column
	 * @param across merge across each row.
	 * @return {@link Ref} array where the affected formula cell references in index 1 and to be evaluated formula cell references in index 0.
	 */
	@SuppressWarnings("unchecked")
	public static ChangeInfo merge(XSheet sheet, int tRow, int lCol, int bRow, int rCol, boolean across) {
		if (across) {
			final Set<Ref> toEval = new HashSet<Ref>();
			final Set<Ref> affected = new HashSet<Ref>();
			final List<MergeChange> changes = new ArrayList<MergeChange>();
			for(int r = tRow; r <= bRow; ++r) {
				final ChangeInfo info = merge0(sheet, r, lCol, r, rCol);
				changes.addAll(info.getMergeChanges());
				toEval.addAll(info.getToEval());
				affected.addAll(info.getAffected());
			}
			return new ChangeInfo(toEval, affected, changes);
		} else {
			return merge0(sheet, tRow, lCol, bRow, rCol);
		}
	}
	
	private static ChangeInfo merge0(XSheet sheet, int tRow, int lCol, int bRow, int rCol) {
		final List<MergeChange> changes = new ArrayList<MergeChange>();
		final Set<Ref> all = new HashSet<Ref>();
		final Set<Ref> last = new HashSet<Ref>();
		//find the left most non-blank cell.
		Cell target = null;
		for(int r = tRow; target == null && r <= bRow; ++r) {
			for(int c = lCol; c <= rCol; ++c) {
				final Cell cell = BookHelper.getCell(sheet, r, c);
				if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK) {
					target = cell;
					break;
				}
			}
		}
		
		CellStyle style = null;
		if (target != null) { //found the target
			final int tgtRow = target.getRowIndex();
			final int tgtCol = target.getColumnIndex();
			final int nRow = tRow - tgtRow;
			final int nCol = lCol - tgtCol;
			if (nRow != 0 || nCol != 0) { //if target not the left-top one, move to left-top
				final ChangeInfo info = BookHelper.moveRange(sheet, tgtRow, tgtCol, tgtRow, tgtCol, nRow, nCol);
				if (info != null) {
					changes.addAll(info.getMergeChanges());
					last.addAll(info.getToEval());
					all.addAll(info.getAffected());
				}
			}
			final CellStyle source = target.getCellStyle();
			style = source.getIndex() == 0 ? null : sheet.getWorkbook().createCellStyle();
			if (style != null) {
				style.cloneStyleFrom(source);
				style.setBorderLeft(CellStyle.BORDER_NONE);
				style.setBorderTop(CellStyle.BORDER_NONE);
				style.setBorderRight(CellStyle.BORDER_NONE);
				style.setBorderBottom(CellStyle.BORDER_NONE);
				target.setCellStyle(style); //set all cell in the merged range to CellStyle of the target minus border
			}
			//1st row (exclude 1st cell)
			for (int c = lCol + 1; c <= rCol; ++c) {
				final Cell cell = getOrCreateCell(sheet, tRow, c);
				cell.setCellStyle(style); //set all cell in the merged range to CellStyle of the target minus border
				final Set<Ref>[] refs = BookHelper.setCellValue(cell, (RichTextString) null);
				if (refs != null) {
					last.addAll(refs[0]);
					all.addAll(refs[1]);
				}
			}
			//2nd row and after
			for(int r = tRow+1; r <= bRow; ++r) {
				for(int c = lCol; c <= rCol; ++c) {
					final Cell cell = getOrCreateCell(sheet, r, c);
					cell.setCellStyle(style); //set all cell in the merged range to CellStyle of the target minus border
					final Set<Ref>[] refs = BookHelper.setCellValue(cell, (RichTextString) null);
					if (refs != null) {
						last.addAll(refs[0]);
						all.addAll(refs[1]);
					}
				}
			}
		}
		
		sheet.addMergedRegion(new CellRangeAddress(tRow, bRow, lCol, rCol));
		final Ref mergeArea = new AreaRefImpl(tRow, lCol, bRow, rCol, BookHelper.getRefSheet((XBook)sheet.getWorkbook(), sheet)); 
		all.add(mergeArea);
		changes.add(new MergeChange(null, mergeArea));
		
		return new ChangeInfo(last, all, changes);
	}	

}
