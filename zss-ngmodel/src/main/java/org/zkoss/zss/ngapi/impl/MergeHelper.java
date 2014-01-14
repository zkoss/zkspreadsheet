package org.zkoss.zss.ngapi.impl;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.InvalidateModelOpException;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NSheet;

public class MergeHelper extends RangeHelperBase{

	public MergeHelper(NRange range) {
		super(range);
	}
	
	public ChangeInfo unmerge(boolean overlapped) {
		int tRow = getRow();
		int lCol = getColumn();
		int bRow = getLastColumn();
		int rCol = getLastColumn();
		
//		final RefSheet refSheet = BookHelper.getRefSheet((XBook)sheet.getWorkbook(), sheet);
		final List<MergeChange> changes = new ArrayList<MergeChange>(); 
		for(int j = sheet.getNumOfMergedRegion() - 1; j >= 0; --j) {
        	final CellRegion merged = sheet.getMergedRegion(j);
        	
        	final int firstCol = merged.getColumn();
        	final int lastCol = merged.getLastColumn();
        	final int firstRow = merged.getRow();
        	final int lastRow = merged.getLastRow();
        	
        	// ZSS-395 unmerge when any cell overlap with merged region
        	// ZSS-412 use a flag to decide to check overlap or not.
        	if( (overlapped && overlap(firstRow, firstCol, lastRow, lastCol, tRow, lCol, bRow, rCol)) || 
        			(!overlapped && contain(tRow, lCol, bRow, rCol,firstRow, firstCol, lastRow, lastCol)) ) {
				changes.add(new MergeChange(merged,null));
				sheet.removeMergedRegion(merged);
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
	public ChangeInfo merge(boolean across) {
		int tRow = range.getRow();
		int lCol = range.getColumn();
		int bRow = range.getLastColumn();
		int rCol = range.getLastColumn();
		
		if(sheet.getOverlapsMergedRegions(new CellRegion(tRow,lCol,bRow,rCol))!=null){
			throw new InvalidateModelOpException("can merge an overlapped region, unmerge it first");
		}
		
		
		if (across) {
			final Set<CellRegion> toEval = new HashSet<CellRegion>();
			final Set<CellRegion> affected = new HashSet<CellRegion>();
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
	
	private ChangeInfo merge0(NSheet sheet, int tRow, int lCol, int bRow, int rCol) {
		final List<MergeChange> changes = new ArrayList<MergeChange>();
		final Set<CellRegion> all = new HashSet<CellRegion>();
		final Set<CellRegion> last = new HashSet<CellRegion>();
		//find the left most non-blank cell.
		NCell target = null;
		for(int r = tRow; target == null && r <= bRow; ++r) {
			for(int c = lCol; c <= rCol; ++c) {
				final NCell cell = sheet.getCell(r, c);
				if (!isBlank(cell)) {
					target = cell;
					break;
				}
			}
		}
		
		NCellStyle style = null;
		if (target != null) { //found the target
			final int tgtRow = target.getRowIndex();
			final int tgtCol = target.getColumnIndex();
			final int nRow = tRow - tgtRow;
			final int nCol = lCol - tgtCol;
			if (nRow != 0 || nCol != 0) { //if target not the left-top one, move to left-top
				//TODO zss 3.5 check and implement this
//				final ChangeInfo info = BookHelper.moveRange(sheet, tgtRow, tgtCol, tgtRow, tgtCol, nRow, nCol);
//				if (info != null) {
//					changes.addAll(info.getMergeChanges());
//					last.addAll(info.getToEval());
//					all.addAll(info.getAffected());
//				}
			}
			final NCellStyle source = target.getCellStyle();
			style = source.equals(sheet.getBook().getDefaultCellStyle()) ? null : sheet.getBook().createCellStyle(source,true);
			if (style != null) {
//				style.cloneStyleFrom(source);
				style.setBorderLeft(BorderType.NONE);
				style.setBorderTop(BorderType.NONE);
				style.setBorderRight(BorderType.NONE);
				style.setBorderBottom(BorderType.NONE);
				target.setCellStyle(style); //set all cell in the merged range to CellStyle of the target minus border
			}
			//1st row (exclude 1st cell)
			for (int c = lCol + 1; c <= rCol; ++c) {
				final NCell cell = sheet.getCell(tRow, c);
				cell.setCellStyle(style); //set all cell in the merged range to CellStyle of the target minus border
				cell.setValue(null);
//				final Set<Ref>[] refs = BookHelper.setCellValue(cell, (RichTextString) null);
//				if (refs != null) {
//					last.addAll(refs[0]);
//					all.addAll(refs[1]);
//				}
			}
			//2nd row and after
			for(int r = tRow+1; r <= bRow; ++r) {
				for(int c = lCol; c <= rCol; ++c) {
					final NCell cell = sheet.getCell(r, c);
					cell.setCellStyle(style); //set all cell in the merged range to CellStyle of the target minus border
					cell.setValue(null);
//					final Set<Ref>[] refs = BookHelper.setCellValue(cell, (RichTextString) null);
//					if (refs != null) {
//						last.addAll(refs[0]);
//						all.addAll(refs[1]);
//					}
				}
			}
		}
		final CellRegion mergeArea = new CellRegion(tRow, bRow, lCol, rCol);
		sheet.addMergedRegion(mergeArea);
//		final Ref mergeArea = new AreaRefImpl(tRow, lCol, bRow, rCol, BookHelper.getRefSheet((XBook)sheet.getWorkbook(), sheet)); 
		all.add(mergeArea);//should update the cell in the merge area.
		changes.add(new MergeChange(null, mergeArea));
		
		return new ChangeInfo(last, all, changes);
	}	

}
