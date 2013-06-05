/* RefSheetImpl.java

	Purpose:
		
	Description:
		
	History:
		Mar 6, 2010 3:39:02 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.engine.impl;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.Map;
import java.util.Set;

import org.zkoss.poi.ss.util.CellReference;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.engine.RefBook;
import org.zkoss.zss.engine.RefSheet;

/**
 * Implementation of {@link RefSheet} for hit testing among reference areas{@link Ref} 
 * and changed cell.
 * <p>This implementation uses the row and column indexes of the two corners of 
 * the reference area as the keys to locate the corresponding {@link Ref}(the 
 * reference area). Each key is ordered in separate list and can be found by 
 * binary search.</p>
 * 
 * <p>When we want to do hit test, given a point of specified row0 and column0(a cell reference), 
 * we say the point hit a reference area if the cell is inside the reference area. That is,
 * row0 >= topRow &amp;&amp; col0 >= leftColumn &amp;&amp; row0 <= bottomRow &amp;&amp; col0 <= rightColumn. 
 * Store index keys in order makes area scanning more efficient because we can avoid scanning those 
 * areas not possible to be the candidates. e.g. For areas' top row scanning, we
 * can scan in area top row list in order one by one. Whenever row0 is less than the 
 * reference area's top row, we can stop scanning immediately since index keys are 
 * stored in order that further reference area's top row will be only bigger than the 
 * previous reference area' top row. That is, any further scan will be only fail 
 * so we can just stop. For area's bottom row scanning, the same algorithm is applied. 
 * The only difference is that this time we shall scan in reverse order and check 
 * whether row0 is greater than the area's bottom row index. And the same algorithm is 
 * applied to two area columns scanning. 
 *    
 * @author henrichen
 *
 */
public class RefSheetImpl implements RefSheet {
	private final Map<RefAddr, Ref> _ltrbIndex; //index of Ref left/top/right/bottom 
	private final IndexArrayList _tRowIndex; //index of Ref top row
	private final IndexArrayList _bRowIndex; //index of Ref bottom row
	private final IndexArrayList _lColIndex; //index of Ref left column
	private final IndexArrayList _rColIndex; //index of Ref right column
	private final RefBook _ownerBook;
	private String _sheetName;
	private final Set<Ref> _indirectDependentRefs; //formula reference contains indirect function(always hit)
	
	public RefSheetImpl(RefBook ownerBook, String sheetName) {
		_ownerBook = ownerBook;
		_sheetName = sheetName;
		_ltrbIndex = new HashMap<RefAddr, Ref>();
		_tRowIndex = new IndexArrayList();
		_bRowIndex = new IndexArrayList();
		_lColIndex = new IndexArrayList();
		_rColIndex = new IndexArrayList();
		_indirectDependentRefs = new HashSet<Ref>(16);
	}
	
	/*package*/void setSheetName(String newsheetname) {
		_sheetName = newsheetname;
	}
	
	@Override
	public String getSheetName() {
		return _sheetName;
	}
	
	@Override
	public RefBook getOwnerBook() {
		return _ownerBook;
	}
	
	/* (non-Javadoc)
	 * @see org.zkoss.zss.engine.RefMatrix#addRef(int, int, int, int)
	 */
	@Override
	public Ref getOrCreateRef(int tRow, int lCol, int bRow, int rCol) {
		final RefAddr addr = new RefAddr(tRow, lCol, bRow, rCol); 
		final Ref ref = _ltrbIndex.get(addr);
		if (ref != null) {
			return ref;
		}
		//a new reference
		//add the new Ref into the RefIndex
		final Ref candidateRef = tRow == bRow && lCol == rCol ?
			new CellRefImpl(tRow, lCol, this) : new AreaRefImpl(tRow, lCol, bRow, rCol, this);
		
		//update the ltrb index
		_ltrbIndex.put(addr, candidateRef);
		
		final Indexable tri = new RefIndex(tRow);
		final Indexable bri = new RefIndex(bRow);
		final Indexable lci = new RefIndex(lCol);
		final Indexable rci = new RefIndex(rCol);
		
		//update the indexes
		final RefIndex tRowRefIndex = (RefIndex)_tRowIndex.getOrAddIndexable(tri);
		final RefIndex bRowRefIndex = (RefIndex)_bRowIndex.getOrAddIndexable(bri);
		final RefIndex lColRefIndex = (RefIndex)_lColIndex.getOrAddIndexable(lci);
		final RefIndex rColRefIndex = (RefIndex)_rColIndex.getOrAddIndexable(rci);
		
		tRowRefIndex.addRef(candidateRef);
		bRowRefIndex.addRef(candidateRef);
		lColRefIndex.addRef(candidateRef);
		rColRefIndex.addRef(candidateRef);
		
		return candidateRef; 
	}

	@Override
	public Ref getRef(int tRow, int lCol, int bRow, int rCol) {
		final RefAddr addr = new RefAddr(tRow, lCol, bRow, rCol); 
		return _ltrbIndex.get(addr);
	}
	
	/* (non-Javadoc)
	 * @see org.zkoss.zss.engine.RefMatrix#getHitRefs(int, int)
	 */
	@SuppressWarnings("unchecked")
	@Override
	public Set<Ref> getHitRefs(int row, int col) {
		final Set<Ref> tRowHits = new HashSet<Ref>();
		final Set<Ref> bRowHits = new HashSet<Ref>();
		final Set<Ref> lColHits = new HashSet<Ref>();
		final Set<Ref> rColHits = new HashSet<Ref>();
		
		//tRow hits
		for(int j=0, len = _tRowIndex.size(); j < len; ++j) {
			final RefIndex tRowRefIndex = (RefIndex) _tRowIndex.get(j);  
			final int tRow = tRowRefIndex.getIndex();
			if (tRow > row) break; //no more
			tRowHits.addAll(tRowRefIndex.getRefs());
		}
		
		int size = tRowHits.size();
		if (size == 0) //special case 
			return _indirectDependentRefs;
		int min = size;
		Set<Ref> hits = tRowHits;
		
		//bRow hits
		for(int j=_bRowIndex.size() - 1; j >= 0; --j) {
			final RefIndex bRowRefIndex = (RefIndex) _bRowIndex.get(j);  
			final int bRow = bRowRefIndex.getIndex();
			if (bRow < row) break; //no more
			bRowHits.addAll(bRowRefIndex.getRefs());
		}
		
		size = bRowHits.size();
		if (size == 0) //special case 
			return _indirectDependentRefs;
		if (min > size) {
			min = size;
			hits = bRowHits;
		}
			
		//lCol hits
		for(int j=0, len = _lColIndex.size(); j < len; ++j) {
			final RefIndex lColRefIndex = (RefIndex) _lColIndex.get(j);  
			final int lCol = lColRefIndex.getIndex();
			if (lCol > col) break; //no more
			lColHits.addAll(lColRefIndex.getRefs());
		}
		
		size = lColHits.size();
		if (size == 0) //special case
			return _indirectDependentRefs;
		if (min > size) {
			min = size;
			hits = lColHits;
		}

		//rCol hits
		for(int j=_rColIndex.size() - 1; j >= 0; --j) {
			final RefIndex rColRefIndex = (RefIndex) _rColIndex.get(j);  
			final int rCol = rColRefIndex.getIndex();
			if (rCol < col) break; //no more
			rColHits.addAll(rColRefIndex.getRefs());
		}
		size = rColHits.size();
		if (size == 0) //special case
			return _indirectDependentRefs;
		if (min > size) {
			min = size;
			hits = rColHits;
		}
		
		final Set<Ref>[] allHits = new Set[]{tRowHits, bRowHits, lColHits, rColHits};
		for(int j=3; j >= 0; --j) {
			Set<Ref> c = allHits[j]; 
			if (hits == c) //yes ==
				continue;
			hits.retainAll(c);
			if (hits.isEmpty()) //special case
				return _indirectDependentRefs;
		}
		
		//add indirectDependentRefs
		hits.addAll(_indirectDependentRefs);
		
		return hits;
	}

	/* (non-Javadoc)
	 * @see org.zkoss.zss.engine.RefMatrix#removeRef(int, int, int, int)
	 */
	@Override
	public Ref removeRef(int tRow, int lCol, int bRow, int rCol) {
		final RefAddr addr = new RefAddr(tRow, lCol, bRow, rCol); 
		final Ref ref = _ltrbIndex.get(addr);
		if (ref == null)
			return null;
		removeRefDirectly(ref);
		
		return ref;
	}
	
	private void removeRefDirectly(Ref ref) {
		final int tRow = ref.getTopRow();
		final int lCol = ref.getLeftCol();
		final int bRow = ref.getBottomRow();
		final int rCol = ref.getRightCol();
			
		final RefAddr addr = new RefAddr(tRow, lCol, bRow, rCol);
		
		//update ltrb index
		_ltrbIndex.remove(addr);
		
		//update 4 indexes
		remove4Indexes(ref);
		
		//update _indirectDependentRefs
		_indirectDependentRefs.remove(ref);
	}
		
	private void remove4Indexes(Ref ref) {
		final int tRow = ref.getTopRow();
		final int lCol = ref.getLeftCol();
		final int bRow = ref.getBottomRow();
		final int rCol = ref.getRightCol();
		
		//update indexes
		final RefIndex tRowRefIndex = (RefIndex) _tRowIndex.getIndexable(tRow);
		if (tRowRefIndex != null) {
			tRowRefIndex.getRefs().remove(ref);
			if (tRowRefIndex.getRefs().isEmpty()) {
				_tRowIndex.removeIndexable(tRow);
			}
		}
		
		final RefIndex bRowRefIndex = (RefIndex) _bRowIndex.getIndexable(bRow);
		if (bRowRefIndex != null) {
			bRowRefIndex.getRefs().remove(ref);
			if (bRowRefIndex.getRefs().isEmpty()) {
				_bRowIndex.removeIndexable(bRow);
			}
		}
		
		final RefIndex lColRefIndex = (RefIndex) _lColIndex.getIndexable(lCol);
		if (lColRefIndex != null) {
			lColRefIndex.getRefs().remove(ref);
			if (lColRefIndex.getRefs().isEmpty()) {
				_lColIndex.removeIndexable(lCol);
			}
		}
		
		final RefIndex rColRefIndex = (RefIndex) _rColIndex.getIndexable(rCol);
		if (rColRefIndex != null) {
			rColRefIndex.getRefs().remove(ref);
			if (rColRefIndex.getRefs().isEmpty()) {
				_rColIndex.removeIndexable(rCol);
			}
		}
	}

	@Override
	public void addDependency(int srcRow, int srcCol, RefSheet sheet,
		int tRow, int lCol, int bRow, int rCol) {
		final CellRefImpl srcRef = (CellRefImpl) 
			getOrCreateRef(srcRow, srcCol, srcRow, srcCol);
		if (sheet == null) {
			sheet = this;
		}
		srcRef.addPrecedent(sheet, tRow, lCol, bRow, rCol);
	}

	@Override
	public void removeDependency(int srcRow, int srcCol, RefSheet sheet, 
		int tRow, int lCol, int bRow, int rCol) {
		final CellRefImpl srcRef = 
			(CellRefImpl) getRef(srcRow, srcCol, srcRow, srcCol);
		if (srcRef != null) {
			if (sheet == null) {
				sheet = this;
			}
			final Ref precedent = srcRef.removePrecedent(sheet, tRow, lCol, bRow, rCol);
			if (precedent.getDependents().isEmpty() && precedent.getPrecedents().isEmpty()) {
				sheet.removeRef(tRow, lCol, bRow, rCol);
			}
			if (srcRef.getDependents().isEmpty() && srcRef.getPrecedents().isEmpty()) {
				removeRef(srcRow, srcCol, srcRow, srcCol);
			}
		}
	}
	@Override
	public void addDependency(int srcRow, int srcCol, String name) {
		final CellRefImpl srcRef = (CellRefImpl) getOrCreateRef(srcRow, srcCol, srcRow, srcCol);
		srcRef.addPrecedent(name);
	}

	@Override
	public void removeDependency(int srcRow, int srcCol, String name) {
		final CellRefImpl srcRef = (CellRefImpl) getRef(srcRow, srcCol, srcRow, srcCol);
		if (srcRef != null) {
			final Ref precedent = srcRef.removePrecedent(name);
			
			if (precedent.getDependents().isEmpty() && precedent.getPrecedents().isEmpty()) {
				_ownerBook.removeVariableRef(name);
			}
			if (srcRef.getDependents().isEmpty() && srcRef.getPrecedents().isEmpty()) {
				removeRef(srcRow, srcCol, srcRow, srcCol);
			}
		}
	}
	
	@Override
	public Set<Ref> getDirectPrecedents(int row, int col) {
		final CellRefImpl srcRef = (CellRefImpl) getRef(row, col, row, col);
		return srcRef != null ? srcRef.getPrecedents() : null;
	}
	
	@Override
	public Set<Ref> getAllPrecedents(int row, int col) {
		final Set<Ref> all = new HashSet<Ref>();
		getAllPrecedents0(row, col, all);
		return all;
	}
	
	private void getAllPrecedents0(int row, int col, Set<Ref> all) {
		final CellRefImpl srcRef = (CellRefImpl) getRef(row, col, row, col);
		if (srcRef != null && !all.contains(srcRef)) {
			final Set<Ref> precedents = srcRef.getPrecedents();
			all.addAll(precedents);
			for (Ref ref : precedents) {
				final int row1 = ref.getTopRow();
				final int row2 = ref.getBottomRow();
				final int col1 = ref.getLeftCol();
				final int col2 = ref.getRightCol();
				final RefSheetImpl refSheet = (RefSheetImpl)ref.getOwnerSheet();
				for (int r = row1; r <= row2; ++r) {
					for (int c = col1; c <= col2; ++c) {
						refSheet.getAllPrecedents0(r, c, all); //recursive
					}
				}
			}
		}
	}
	
	@Override
	public Set<Ref> getDirectDependents(int row, int col) {
		return DependencyTrackerHelper.getDirectDependents(this, row, col);
	}
		
	@Override
	public Set<Ref> getAllDependents(int row, int col) {
		final Set<Ref> all = new HashSet<Ref>();
		DependencyTrackerHelper.getBothDependents(this, row, col, all, null);
		return all;
	}
	
	@Override
	public Set<Ref> getLastDependents(int row, int col) {
		final Set<Ref> last = new HashSet<Ref>();
		final Set<Ref> all = new HashSet<Ref>();
		DependencyTrackerHelper.getBothDependents(this, row, col, all, last);
		return last;
	}
	
	@SuppressWarnings("unchecked")
	@Override
	public Set<Ref>[] getBothDependents(int row, int col) {
		final Set<Ref> last = new HashSet<Ref>();
		final Set<Ref> all = new HashSet<Ref>();
		DependencyTrackerHelper.getBothDependents(this, row, col, all, last);
		return (Set<Ref>[]) new Set[] {last, all};
	}
	
	@Override
	public Set<Ref>[] insertRows(int startRow, int num) {
		if (num <= 0) {
			throw new UiException("Cannot insert a negative number of rows: "+num);
		}
		final Set<Ref> hits = new HashSet<Ref>();
		
		//bRow hits
		for(int j=_bRowIndex.size() - 1; j >= 0; --j) {
			final RefIndex bRowRefIndex = (RefIndex) _bRowIndex.get(j);  
			final int bRow = bRowRefIndex.getIndex();
			if (bRow < startRow) break; //no more
			hits.addAll(bRowRefIndex.getRefs());
		}
		
		//remove hits from the ltrbIndex(before adjust row
		removeFromLtrbIndex(hits);

		final int maxrow = _ownerBook.getMaxrow();
		//adjust Ref bRow index
		for(int j=_bRowIndex.size() - 1; j >= 0; --j) {
			final RefIndex bRowRefIndex = (RefIndex) _bRowIndex.get(j);  
			int bRow = bRowRefIndex.getIndex();
			if (bRow < startRow) break; //no more
			bRow += num;
			final Set<Ref> bRowRefs = bRowRefIndex.getRefs();
			if (bRow > maxrow) {//out of maximum bound
				_bRowIndex.remove(j); //remove the RefIndex of this bottom row
				bRow = maxrow;
			} else {
				bRowRefIndex.setIndex(bRow); //assign a new index
			}
			for(Ref ref : bRowRefs) {
				ref.setBottomRow(bRow);
			}
		}
		
		//adjust Ref tRow index
		final Set<Ref> removeHits = new HashSet<Ref>();
		for(int j=_tRowIndex.size() - 1; j >= 0; --j) {
			final RefIndex tRowRefIndex = (RefIndex) _tRowIndex.get(j);  
			int tRow = tRowRefIndex.getIndex();
			if (tRow < startRow) break; //no more
			tRow += num;
			final Set<Ref> tRowRefs = tRowRefIndex.getRefs();
			if (tRow > maxrow) {//out of maximum bound, #REF! case
				_bRowIndex.remove(j); //remove the RefIndex of this top row
				removeHits.addAll(tRowRefs);
			} else {
				tRowRefIndex.setIndex(tRow); //assign a new index
				for(Ref ref : tRowRefs) {
					ref.setTopRow(tRow);
				}
			}
		}

		//add/merge back the adjusted Refs to ltrb index
		addOrMergeBackLtrbIndex(hits);

		//clear precedents link
		clearRemoveRefs(removeHits, hits);
		
		return getBothDependents(removeHits, hits);
	}

	//clear precendents link (keep dependents link)
	private void clearRemoveRefs(Set<Ref> removeHits, Set<Ref> hits) {
		for (Ref ref : removeHits) {
			//remove from 4 indexes
			remove4Indexes(ref);
			//clear precedents
			ref.removeAllPrecedents();
			final Set<Ref> dependents = ref.getDependents();
			if (dependents.isEmpty()) {
				hits.remove(ref);
			} else {
				for (Ref dependent : dependents) {
					dependent.getPrecedents().remove(ref);
					//orphan Ref, remove self indexes
					if (dependent.getPrecedents().isEmpty() && dependent.getDependents().isEmpty() && !removeHits.contains(dependent)) {
						((RefSheetImpl)dependent.getOwnerSheet()).removeRefDirectly(dependent);
					}
				}
			}
		}
	}
	
	@Override
	public Set<Ref>[] deleteRows(int startRow, int num) {
		if (num <= 0) {
			throw new UiException("Cannot remove a negative number of rows: "+num);
		}
		final Set<Ref> hits = new HashSet<Ref>();
		
		//bRow hits
		int jb = _bRowIndex.size() - 1;
		for(; jb >= 0; --jb) {
			final RefIndex bRowRefIndex = (RefIndex) _bRowIndex.get(jb);  
			final int bRow = bRowRefIndex.getIndex();
			if (bRow < startRow) break; //no more
			hits.addAll(bRowRefIndex.getRefs());
		}

		//remove hits from the ltrbIndex(before adjust row)
		removeFromLtrbIndex(hits);
		
		//adjust Ref bRow index
		for(int j = jb+1, len = _bRowIndex.size(); j < len; ++j) {
			final RefIndex bRowRefIndex = (RefIndex) _bRowIndex.get(j);
			int bRow = bRowRefIndex.getIndex();
			int min = bRow - startRow + 1;
			if (min > num) min = num;
			if (min == 0) {
				continue; //skip, nothing to adjust
			}
			bRow -= min;
			bRowRefIndex.setIndex(bRow);
			_bRowIndex.remove(j);
			final RefIndex newRefIndex = (RefIndex) _bRowIndex.getOrAddIndexable(bRowRefIndex); //try to move the bRowRefIndex
			boolean copy = false;
			if (System.identityHashCode(newRefIndex) != System.identityHashCode(bRowRefIndex)) { //no way to move, use copy
				--j;
				--len;
				copy = true;
			}
			final Set<Ref> bRowRefs = bRowRefIndex.getRefs();
			for(Ref ref : bRowRefs) {
				ref.setBottomRow(bRow);
			}
			if (copy) {
				newRefIndex.getRefs().addAll(bRowRefs); 
			}
		}
		
		final Set<Ref> removeHits = new HashSet<Ref>();
		//adjust Ref tRow index
		for(int j = _tRowIndex.getListIndex(startRow), len = _tRowIndex.size(); j < len; ++j) {
			final RefIndex tRowRefIndex = (RefIndex) _tRowIndex.get(j);
			int tRow = tRowRefIndex.getIndex();
			int min = tRow - startRow;
			if (min > num) min = num;
			if (min != 0) {
				tRow -= min;
				tRowRefIndex.setIndex(tRow);
				_tRowIndex.remove(j);
			}
			final RefIndex newRefIndex = (RefIndex) _tRowIndex.getOrAddIndexable(tRowRefIndex); //try to move the tRowRefIndex
			boolean copy = false;
			if (System.identityHashCode(newRefIndex) != System.identityHashCode(tRowRefIndex)) { //no way to move, use copy
				--j;
				--len;
				copy = true;
			}
			final Set<Ref> tRowRefs = tRowRefIndex.getRefs();
			if (copy) {
				final Set<Ref> newRefs = newRefIndex.getRefs(); 
				for(Ref ref : tRowRefs) {
					if (ref.getBottomRow() < tRow) { //shall be removed since bRow < tRow
						removeHits.add(ref);
					} else {
						ref.setTopRow(tRow);
						newRefs.add(ref);
					}
				}
			} else {
				for(Ref ref : tRowRefs) {
					if (ref.getBottomRow() < tRow) { //shall be removed since bRow < tRow
						removeHits.add(ref);
					} else if (min != 0) {
						ref.setTopRow(tRow);
					}
				}
			}
		}
		
		//add back the adjusted Refs to ltrb index
		addOrMergeBackLtrbIndex(hits);

		//clear removed refs (Keep the dependents link)
		clearRemoveRefs(removeHits, hits);
		
		return getBothDependents(removeHits, hits);
	}

	/**
	 * 
	 * @param removeHits Refs to be removed
	 * @param precedents Refs whose contents is changed 
	 * @return [0]: toEval, [1]: affected
	 */
	@SuppressWarnings("unchecked")
	private Set<Ref>[] getBothDependents(Set<Ref> removeHits, Set<Ref> precedents) {
		final Set<Ref> all = new HashSet<Ref>();
		final Set<Ref> last = new HashSet<Ref>();
		final Set<Ref> dependents = DependencyTrackerHelper.getDirectDependents(precedents);
		DependencyTrackerHelper.getBothDependents(dependents, all, last);
		all.removeAll(removeHits);
		last.removeAll(removeHits);
		return (Set<Ref>[]) new Set[] {last, all};
	}
	
	@Override
	public Set<Ref>[] insertColumns(int startCol, int num) {
		if (num <= 0) {
			throw new UiException("Cannot insert a negative number of columns: "+num);
		}
		final Set<Ref> hits = new HashSet<Ref>();
		
		//rCol hits
		for(int j=_rColIndex.size() - 1; j >= 0; --j) {
			final RefIndex rColRefIndex = (RefIndex) _rColIndex.get(j);  
			final int rCol = rColRefIndex.getIndex();
			if (rCol < startCol) break; //no more
			hits.addAll(rColRefIndex.getRefs());
		}
		
		//remove hits from the ltrbIndex(before adjust column)
		removeFromLtrbIndex(hits);

		final int maxcol = _ownerBook.getMaxcol();
		//adjust Ref rCol index
		for(int j=_rColIndex.size() - 1; j >= 0; --j) {
			final RefIndex rColRefIndex = (RefIndex) _rColIndex.get(j);  
			int rCol = rColRefIndex.getIndex();
			if (rCol < startCol) break; //no more
			rCol += num;
			final Set<Ref> rColRefs = rColRefIndex.getRefs();
			if (rCol > maxcol) {//out of maximum bound
				_rColIndex.remove(j); //remove the RefIndex of this right column
				rCol = maxcol;
			} else {
				rColRefIndex.setIndex(rCol); //assign a new index
			}
			for(Ref ref : rColRefs) {
				ref.setRightCol(rCol);
			}
		}
		
		//adjust Ref lCol index
		final Set<Ref> removeHits = new HashSet<Ref>();
		for(int j=_lColIndex.size() - 1; j >= 0; --j) {
			final RefIndex lColRefIndex = (RefIndex) _lColIndex.get(j);  
			int lCol = lColRefIndex.getIndex();
			if (lCol < startCol) break; //no more
			lCol += num;
			final Set<Ref> lColRefs = lColRefIndex.getRefs();
			if (lCol > maxcol) {//out of maximum bound, #REF! case
				_rColIndex.remove(j); //remove the RefIndex of this left column
				removeHits.addAll(lColRefs);
			} else {
				lColRefIndex.setIndex(lCol); //assign a new index
				for(Ref ref : lColRefs) {
					ref.setLeftCol(lCol);
				}
			}
		}

		//add/merge back the adjusted Refs to ltrb index
		addOrMergeBackLtrbIndex(hits);

		//clear precedents link
		clearRemoveRefs(removeHits, hits);
		
		return getBothDependents(removeHits, hits);
	}
	
	@Override
	public Set<Ref>[] deleteColumns(int startCol, int num) {
		if (num <= 0) {
			throw new UiException("Cannot remove a negative number of columns: "+num);
		}
		final Set<Ref> hits = new HashSet<Ref>();
		
		//rCol hits
		int jb = _rColIndex.size() - 1;
		for(; jb >= 0; --jb) {
			final RefIndex rColRefIndex = (RefIndex) _rColIndex.get(jb);  
			final int rCol = rColRefIndex.getIndex();
			if (rCol < startCol) break; //no more
			hits.addAll(rColRefIndex.getRefs());
		}
		
		//remove hits from the ltrbIndex(before adjust col)
		removeFromLtrbIndex(hits);
		
		//adjust Ref rCol index
		for(int j = jb+1, len = _rColIndex.size(); j < len; ++j) {
			final RefIndex rColRefIndex = (RefIndex) _rColIndex.get(j);
			int rCol = rColRefIndex.getIndex();
			int min = rCol - startCol + 1;
			if (min > num) min = num;
			if (min == 0) {
				continue; //skip, nothing to adjust
			}
			rCol -= min;
			rColRefIndex.setIndex(rCol);
			_rColIndex.remove(j);
			final RefIndex newRefIndex = (RefIndex) _rColIndex.getOrAddIndexable(rColRefIndex); //try to move the rColRefIndex
			boolean copy = false;
			if (System.identityHashCode(newRefIndex) != System.identityHashCode(rColRefIndex)) { //no way to move, use copy
				--j;
				--len;
				copy = true;
			}
			final Set<Ref> rColRefs = rColRefIndex.getRefs();
			for(Ref ref : rColRefs) {
				ref.setRightCol(rCol);
			}
			if (copy) {
				newRefIndex.getRefs().addAll(rColRefs); 
			}
		}
		
		final Set<Ref> removeHits = new HashSet<Ref>();
		//adjust Ref lCol index
		for(int j = _lColIndex.getListIndex(startCol), len = _lColIndex.size(); j < len; ++j) {
			final RefIndex lColRefIndex = (RefIndex) _lColIndex.get(j);
			int lCol = lColRefIndex.getIndex();
			int min = lCol - startCol;
			if (min > num) min = num;
			if (min != 0) {
				lCol -= min;
				lColRefIndex.setIndex(lCol);
				_lColIndex.remove(j);
			}
			final RefIndex newRefIndex = (RefIndex) _lColIndex.getOrAddIndexable(lColRefIndex); //try to move the lColRefIndex
			boolean copy = false;
			if (System.identityHashCode(newRefIndex) != System.identityHashCode(lColRefIndex)) { //no way to move, use copy
				--j;
				--len;
				copy = true;
			}
			final Set<Ref> lColRefs = lColRefIndex.getRefs();
			if (copy) {
				final Set<Ref> newRefs = newRefIndex.getRefs(); 
				for(Ref ref : lColRefs) {
					if (ref.getRightCol() < lCol) { //shall be removed since rCol < lCol
						removeHits.add(ref);
					} else {
						ref.setLeftCol(lCol);
						newRefs.add(ref);
					}
				}
			} else {
				for(Ref ref : lColRefs) {
					if (ref.getRightCol() < lCol) { //shall be removed since rCol < lCol
						removeHits.add(ref);
					} else if (min != 0) {
						ref.setLeftCol(lCol);
					}
				}
			}
		}
		
		//add back the adjusted Refs to ltrb index
		addOrMergeBackLtrbIndex(hits);

		//clear removed refs (Keep the dependents link)
		clearRemoveRefs(removeHits, hits);
		
		return getBothDependents(removeHits, hits);
	}

	private void addOrMergeBackLtrbIndex(Set<Ref> affectRefs) {
		for (Ref ref : affectRefs) {
			addOrMergeBackLtrbIndex(ref);
		}
	}
	
	private void addOrMergeBackLtrbIndex(Ref ref) {
		final RefAddr refAddr = new RefAddr(ref);
		final Ref refX = _ltrbIndex.get(refAddr);
		if (refX == null)
			_ltrbIndex.put(refAddr, ref);
		else {
			final Set<Ref> dependents = ref.getDependents(); 
			refX.getDependents().addAll(dependents);
			for (Ref dependent : dependents) {
				final Set<Ref> depprecendents = dependent.getPrecedents(); 
				depprecendents.remove(ref);
				depprecendents.add(refX);
			}
			final Set<Ref> precedents = ref.getPrecedents();
			refX.getPrecedents().addAll(precedents);
			for (Ref precedent : precedents) {
				final Set<Ref> predependents = precedent.getDependents(); 
				predependents.remove(ref);
				predependents.add(refX);
			}
			dependents.clear();
			precedents.clear();
			remove4Indexes(ref);
		}
	}
	
	private void removeFromLtrbIndex(Set<Ref> refs) {
		for (Ref ref : refs) {
			_ltrbIndex.remove(new RefAddr(ref));
		}
	}
	
	//Return all Refs that intersect with the given area
	private Set<Ref> getIntersectRefs(int startRow, int startCol, int endRow, int endCol) {
		final Set<Ref> hits = new HashSet<Ref>(); //intersect refs
		
		//lCol hits
		for(int jl = 0, len = _lColIndex.size(); jl < len; ++jl) {
			final RefIndex lColRefIndex = (RefIndex) _lColIndex.get(jl);  
			final int lCol = lColRefIndex.getIndex();
			if (lCol > endCol) break; //no more
			hits.addAll(lColRefIndex.getRefs());
		}

		//rCol hits
		Set<Ref> tmpHits = new HashSet<Ref>();
		for(int jr = _rColIndex.size() - 1; jr >= 0; --jr) {
			final RefIndex rColRefIndex = (RefIndex) _rColIndex.get(jr);  
			final int rCol = rColRefIndex.getIndex();
			if (rCol < startCol) break; //no more
			tmpHits.addAll(rColRefIndex.getRefs());
		}
		hits.retainAll(tmpHits);
		
		//bRow hits
		tmpHits = new HashSet<Ref>();
		for(int jb = _bRowIndex.size() - 1; jb >= 0; --jb) {
			final RefIndex bRowRefIndex = (RefIndex) _bRowIndex.get(jb);  
			final int bRow = bRowRefIndex.getIndex();
			if (bRow < startRow) break; //no more
			tmpHits.addAll(bRowRefIndex.getRefs());
		}
		hits.retainAll(tmpHits);
		
		//tRow hits
		tmpHits = new HashSet<Ref>();
		for(int jt = 0, len = _tRowIndex.size(); jt < len; ++jt) {
			final RefIndex tRowRefIndex = (RefIndex) _tRowIndex.get(jt);  
			final int tRow = tRowRefIndex.getIndex();
			if (tRow > endRow) break; //no more
			tmpHits.addAll(tRowRefIndex.getRefs());
		}
		hits.retainAll(tmpHits);
		
		return hits;
		
	}
	
	private Set<Ref> adjustRemoveRefs(Set<Ref> hits, int startRow, int startCol, int endRow, int endCol) {
		//adjust those Ref in removed range
		final Set<Ref> removeHits = new HashSet<Ref>(); //totally removed
		for(Ref ref : hits) { //Refs that intersect with the removing range
			final int lCol = ref.getLeftCol();
			final int rCol = ref.getRightCol();
			final int tRow = ref.getTopRow();
			final int bRow = ref.getBottomRow();
			//final RefAddr refAddr = new RefAddr(tRow, lCol, bRow, rCol);
			if (lCol >= startCol && rCol <= endCol) { //total remove a row of the Ref
				if (tRow < startRow) {
					if (bRow < endRow) { //remove bottom side
						final int newRow = startRow - 1;
						changeRefIndex(_bRowIndex, bRow, newRow, ref);
						ref.setBottomRow(newRow);
					}
				} else {
					if (bRow > endRow) { //remove top side
						if (tRow > startRow) {
							final int newRow = endRow + 1;
							changeRefIndex(_tRowIndex, tRow, newRow, ref);
							ref.setTopRow(newRow);
						}
					} else {
						removeHits.add(ref);
					}
				}
			} else if (tRow >= startRow && bRow <= endRow) { //total remove a column of the Ref
				if (lCol < startCol) {
					if (rCol < endCol) { //remove right side
						final int newCol = startCol - 1;
						changeRefIndex(_rColIndex, rCol, newCol, ref);
						ref.setRightCol(newCol);
					}
				} else {
					if (rCol > endCol) { //remove left side
						if (lCol > startCol) {
							final int newCol = endCol + 1;
							changeRefIndex(_lColIndex, lCol, newCol, ref);
							ref.setLeftCol(newCol);
						}
					}
				}
			}
		}
		return removeHits;
	}
	
	@Override
	public Set<Ref>[] deleteRange(int startRow, int startCol, int endRow, int endCol, boolean horizontal) {
		final Set<Ref> hits = getIntersectRefs(startRow, startCol, endRow, endCol);
		//remove hits from the ltrbIndex(before adjust col and row)
		removeFromLtrbIndex(hits);
		final Set<Ref> removeHits = adjustRemoveRefs(hits, startRow, startCol, endRow, endCol);
		Set<Ref> moveHits = new HashSet<Ref>();
		//adjust those moved because of the remove
		if (horizontal) { //move left
			final int num = endCol - startCol + 1; //how many columns to move
			final int startCol0 = endCol + 1;
			//rCol hits
			for(int jr = _rColIndex.size() - 1; jr >= 0; --jr) {
				final RefIndex rColRefIndex = (RefIndex) _rColIndex.get(jr);  
				final int rCol = rColRefIndex.getIndex();
				if (rCol < startCol0) break; //no more
				moveHits.addAll(rColRefIndex.getRefs());
			}
			
			//bRow hits
			Set<Ref> tmpHits = new HashSet<Ref>();
			for(int jb = _bRowIndex.size() - 1; jb >= 0; --jb) {
				final RefIndex bRowRefIndex = (RefIndex) _bRowIndex.get(jb);  
				final int bRow = bRowRefIndex.getIndex();
				if (bRow < startRow) break; //no more
				tmpHits.addAll(bRowRefIndex.getRefs());
			}
			moveHits.retainAll(tmpHits);
			
			//tRow hits
			for(int jt = 0, len = _tRowIndex.size(); jt < len; ++jt) {
				final RefIndex tRowRefIndex = (RefIndex) _tRowIndex.get(jt);  
				final int tRow = tRowRefIndex.getIndex();
				if (tRow > endRow) break; //no more
				tmpHits.addAll(tRowRefIndex.getRefs());
			}
			moveHits.retainAll(tmpHits);
			
			removeFromLtrbIndex(moveHits);
			hits.addAll(moveHits); //Refs that has changed/removed
			
			for(Ref ref : moveHits) {
				final int lCol = ref.getLeftCol();
				final int rCol = ref.getRightCol();
				final int tRow = ref.getTopRow();
				final int bRow = ref.getBottomRow();
				if (tRow >= startRow && bRow <= endRow) { //total cover
					if (lCol > endCol) {
						final int newCol = lCol - num;
						changeRefIndex(_lColIndex, lCol, newCol, ref);
						ref.setLeftCol(newCol);
					}
					final int newCol = rCol - num;
					changeRefIndex(_rColIndex, rCol, newCol, ref);
					ref.setRightCol(newCol);
				}
			}
		} else { //move up!
			final int num = endRow - startRow + 1; //how many columns to move
			final int startRow0 = endRow + 1;
			//bRow hits
			for(int jb = _bRowIndex.size() - 1; jb >= 0; --jb) {
				final RefIndex bRowRefIndex = (RefIndex) _bRowIndex.get(jb);  
				final int bRow = bRowRefIndex.getIndex();
				if (bRow < startRow0) break; //no more
				moveHits.addAll(bRowRefIndex.getRefs());
			}
			
			//rCol hits
			Set<Ref> tmpHits = new HashSet<Ref>();
			for(int jr = _rColIndex.size() - 1; jr >= 0; --jr) {
				final RefIndex rColRefIndex = (RefIndex) _rColIndex.get(jr);  
				final int rCol = rColRefIndex.getIndex();
				if (rCol < startCol) break; //no more
				tmpHits.addAll(rColRefIndex.getRefs());
			}
			moveHits.retainAll(tmpHits);
			
			//lCol hits
			tmpHits = new HashSet<Ref>();
			for(int jl = 0, len = _lColIndex.size(); jl < len; ++jl) {
				final RefIndex lColRefIndex = (RefIndex) _lColIndex.get(jl);  
				final int lCol = lColRefIndex.getIndex();
				if (lCol > endCol) break; //no more
				tmpHits.addAll(lColRefIndex.getRefs());
			}
			moveHits.retainAll(tmpHits);
			
			removeFromLtrbIndex(moveHits);

			hits.addAll(moveHits); //Refs that has changed/removed
			
			for(Ref ref : moveHits) {
				final int lCol = ref.getLeftCol();
				final int rCol = ref.getRightCol();
				final int tRow = ref.getTopRow();
				final int bRow = ref.getBottomRow();
				if (lCol >= startCol && rCol <= endCol) { //total cover
					if (tRow > endRow) {
						final int newRow = tRow - num; 
						changeRefIndex(_tRowIndex, tRow, newRow, ref);
						ref.setTopRow(newRow);
					}
					final int newRow = bRow - num;
					changeRefIndex(_bRowIndex, bRow, newRow, ref);
					ref.setBottomRow(newRow);
				}
			}
		}
		
		//add back the adjusted Refs to ltrb index
		final Set<Ref> affectHits = new HashSet<Ref>(hits);
		affectHits.removeAll(removeHits);
		addOrMergeBackLtrbIndex(affectHits);

		//clear removed refs (Keep the dependents link)
		clearRemoveRefs(removeHits, hits);
		
		return getBothDependents(removeHits, hits);
	}

	private void changeRefIndex(IndexArrayList indexes, int orgIndex, int newIndex, Ref ref) {
		final RefIndex orgRColRefIndex = (RefIndex) indexes.getIndexable(orgIndex);
		orgRColRefIndex.getRefs().remove(ref);
		if (orgRColRefIndex.getRefs().isEmpty()) {
			indexes.removeIndexable(orgIndex);
		}
		final RefIndex newRColRefIndex = (RefIndex) indexes.getOrAddIndexable(new RefIndex(newIndex));
		newRColRefIndex.addRef(ref);
	}
	
	/**
	 * Index array used to manage object with indexes (such as formula references).
	 * @author henrichen
	 */
	private static class IndexArrayList extends ArrayList<Indexable> {
		private static final long serialVersionUID = 201003061603L;

		/**
		 * Returns the indexable in the specified index if exists; otherwise adds  
		 * the specified indexable in the indexed place and return it back.  
		 * @param indexable the indexable might be added if no indexable in the specified index.
		 * @return the indexable in the indexed place.
		 */
		public Indexable getOrAddIndexable(Indexable indexable) {
			final int j = Collections.binarySearch(this, indexable);
			if (j < 0) { //not there, add into list
				add(-1 - j, indexable);
				return indexable;
			} else  //found, use the one in list
				return get(j); 
		}
		
		/**
		 * Returns the associated list index at or next to(if not exist the indexable) the specified indexable index.
		 * @param index the indexable index.
		 * @return the associated list index at or next to(if not exist the indexable) the specified indexable index.
		 */
		public int getListIndex(int index) {
			final int j = Collections.binarySearch(this, new DummyIndexable(index));
			return j < 0 ? -1 - j : j;
		}
		
		/**
		 * Returns the indexable in the specified index.  
		 * @param index the indexable index.
		 * @return the indexable in the indexed place.
		 */
		public Indexable getIndexable(int index) {
			final int j = Collections.binarySearch(this, new DummyIndexable(index));
			return j < 0 ? null : get(j);
		}
		
		/**
		 * Remove the indexable in the specified indexable index and return it; return null if not exist.
		 * @param index the indexable index
 		 * @return the removed indexable
		 */
		public Indexable removeIndexable(int index) {
			final int j = Collections.binarySearch(this, new DummyIndexable(index));
			return j < 0 ? null : remove(j); 
		}
	}
	
	private static class DummyIndexable implements Indexable {
		private int _dummyIndex;
		
		public DummyIndexable(int index) {
			_dummyIndex = index;
		}
		
		public int getIndex() {
			return _dummyIndex;
		}

		@Override
		public int compareTo(Indexable o) {
			return this.getIndex() - o.getIndex();
		}
		
		@Override
		public boolean equals(Object other) {
			return this == other
				|| (other instanceof Indexable && compareTo((Indexable) other) == 0);
		}
	}

	private static class RefIndex implements Indexable {
		private int _index;
		private Set<Ref> _refs;

		public RefIndex(int index) {
			_index = index;
			_refs = new HashSet<Ref>(16);
		}
		
		public int getIndex() {
			return _index;
		}
		
		private void setIndex(int index) {
			_index = index;
		}

		private void addRef(Ref ref) {
			_refs.add(ref);
		}
		
		public Set<Ref> getRefs() {
			return _refs;
		}
		
		@Override
		public int compareTo(Indexable o) {
			return this.getIndex() - o.getIndex();
		}
		//--Object--//
		@Override
		public int hashCode() {
			return _index;
		}
		
		@Override
		public boolean equals(Object other) {
			if (other == this) {
				return true;
			}
			if (!(other instanceof RefIndex)) {
				return false;
			}
			return _index == ((RefIndex)other)._index;
		}
	}
	
	@Override
	public Set<Ref>[] insertRange(int startRow, int startCol, int endRow, int endCol, boolean horizontal) {
		final Set<Ref> removeHits = new HashSet<Ref>(); //totally removed
		Set<Ref> hits = new HashSet<Ref>();
		//adjust those moved because of the insert
		if (horizontal) { //move right
			final int num = endCol - startCol + 1; //how many columns to move
			
			//rCol hits
			for(int jr = _rColIndex.size() - 1; jr >= 0; --jr) {
				final RefIndex rColRefIndex = (RefIndex) _rColIndex.get(jr);  
				final int rCol = rColRefIndex.getIndex();
				if (rCol < startCol) break; //no more
				hits.addAll(rColRefIndex.getRefs());
			}
			
			//bRow hits
			Set<Ref> tmpHits = new HashSet<Ref>();
			for(int jb = _bRowIndex.size() - 1; jb >= 0; --jb) {
				final RefIndex bRowRefIndex = (RefIndex) _bRowIndex.get(jb);  
				final int bRow = bRowRefIndex.getIndex();
				if (bRow < startRow) break; //no more
				tmpHits.addAll(bRowRefIndex.getRefs());
			}
			hits.retainAll(tmpHits);
			
			//tRow hits
			tmpHits = new HashSet<Ref>();
			for(int jt = 0, len = _tRowIndex.size(); jt < len; ++jt) {
				final RefIndex tRowRefIndex = (RefIndex) _tRowIndex.get(jt);  
				final int tRow = tRowRefIndex.getIndex();
				if (tRow > endRow) break; //no more
				tmpHits.addAll(tRowRefIndex.getRefs());
			}
			hits.retainAll(tmpHits);
			
			removeFromLtrbIndex(hits);
			
			final int maxcol = _ownerBook.getMaxcol();
			for(Ref ref : hits) {
				final int lCol = ref.getLeftCol();
				final int rCol = ref.getRightCol();
				final int tRow = ref.getTopRow();
				final int bRow = ref.getBottomRow();
				if (tRow >= startRow && bRow <= endRow) { //total cover
					if (lCol >= startCol) {
						final int newCol = lCol + num;
						changeRefIndex(_lColIndex, lCol, newCol, ref);
						ref.setLeftCol(newCol);
					}
					final int newrCol = Math.min(rCol + num, maxcol);
					if (ref.getLeftCol() > newrCol) { //push out of bound
						removeHits.add(ref);
					} else {
						changeRefIndex(_rColIndex, rCol, newrCol, ref);
						ref.setRightCol(newrCol);
					}
				}
			}
		} else { //move down!
			final int num = endRow - startRow + 1; //how many columns to move
			
			//bRow hits
			for(int jb = _bRowIndex.size() - 1; jb >= 0; --jb) {
				final RefIndex bRowRefIndex = (RefIndex) _bRowIndex.get(jb);  
				final int bRow = bRowRefIndex.getIndex();
				if (bRow < startRow) break; //no more
				hits.addAll(bRowRefIndex.getRefs());
			}
			
			//rCol hits
			Set<Ref> tmpHits = new HashSet<Ref>();
			for(int jr = _rColIndex.size() - 1; jr >= 0; --jr) {
				final RefIndex rColRefIndex = (RefIndex) _rColIndex.get(jr);  
				final int rCol = rColRefIndex.getIndex();
				if (rCol < startCol) break; //no more
				tmpHits.addAll(rColRefIndex.getRefs());
			}
			hits.retainAll(tmpHits);
			
			//lCol hits
			tmpHits = new HashSet<Ref>();
			for(int jl = 0, len = _lColIndex.size(); jl < len; ++jl) {
				final RefIndex lColRefIndex = (RefIndex) _lColIndex.get(jl);  
				final int lCol = lColRefIndex.getIndex();
				if (lCol > endCol) break; //no more
				tmpHits.addAll(lColRefIndex.getRefs());
			}
			hits.retainAll(tmpHits);
			
			removeFromLtrbIndex(hits);

			final int maxrow = _ownerBook.getMaxrow();
			for(Ref ref : hits) {
				final int lCol = ref.getLeftCol();
				final int rCol = ref.getRightCol();
				final int tRow = ref.getTopRow();
				final int bRow = ref.getBottomRow();
				if (lCol >= startCol && rCol <= endCol) { //total cover
					if (tRow >= startRow) {
						final int newRow = tRow + num;
						changeRefIndex(_tRowIndex, tRow, newRow, ref);
						ref.setTopRow(newRow);
					}
					final int newbRow = Math.min(bRow + num, maxrow);
					if (ref.getTopRow() > newbRow) { //push out of bound
						removeHits.add(ref);
					} else {
						changeRefIndex(_bRowIndex, bRow, newbRow, ref);
						ref.setBottomRow(newbRow);
					}
				}
			}
		}
		
		//add back the adjusted Refs to ltrb index
		final Set<Ref> affectHits = new HashSet<Ref>(hits);
		affectHits.removeAll(removeHits);
		addOrMergeBackLtrbIndex(affectHits);

		//clear removed refs (Keep the dependents link)
		clearRemoveRefs(removeHits, hits);
		
		return getBothDependents(removeHits, hits);
	}
	
	@Override
	public Set<Ref>[] moveRange(int startRow, int startCol, int endRow, int endCol, int nRow, int nCol) {
		//TODO destination area overlap with source area, special treatment
		final int maxcol = _ownerBook.getMaxcol();
		final int maxrow = _ownerBook.getMaxrow();
		final int dstStartRow = startRow + nRow;
		final int dstEndRow = endRow + nRow;
		final int dstStartCol = startCol + nCol;
		final int dstEndCol = endCol + nCol;
		
		if (dstStartRow < 0 || dstEndRow > maxrow || dstStartCol < 0 || dstEndCol > maxcol) {
			throw new UiException("Move out of bound");
		}
		
		final Set<Ref> srcHits = getIntersectRefs(startRow, startCol, endRow, endCol);
		
		//remove hits from the ltrbIndex(before adjust col and row)
		removeFromLtrbIndex(srcHits);
		
		//destination area (to be removed)
		final Set<Ref> dstHits = getIntersectRefs(dstStartRow, dstStartCol, dstEndRow, dstEndCol);
		dstHits.removeAll(srcHits); //excludes those to be processed in source area
		//remove hits from the ltrbIndex(before adjust col and row)
		removeFromLtrbIndex(dstHits);
		final Set<Ref> removeHits = adjustRemoveRefs(dstHits, dstStartRow, dstStartCol, dstEndRow, dstEndCol);
		
		//adjust those Ref in source range
		for(Ref ref : srcHits) { //Refs that intersect with the removing range
			final int lCol = ref.getLeftCol();
			final int rCol = ref.getRightCol();
			final int tRow = ref.getTopRow();
			final int bRow = ref.getBottomRow();
			//final RefAddr refAddr = new RefAddr(tRow, lCol, bRow, rCol);
			if (lCol >= startCol && rCol <= endCol) { //total cover a row of the Ref
				if (tRow < startRow) {
					if (bRow <= endRow) { //move bottom side
						if (nCol == 0) {
							if (dstEndRow >= tRow) {
								final int newbRow = Math.max(bRow + nRow, startRow - 1);
								changeRefIndex(_bRowIndex, bRow, newbRow, ref);
								ref.setBottomRow(newbRow);
								if (dstStartRow < tRow) {
									changeRefIndex(_tRowIndex, tRow, dstStartRow, ref);
									ref.setTopRow(dstStartRow);
								}
							}
						}
					}
				} else {
					if (bRow > endRow) { //move top side
						if (nCol == 0) {
							if (dstStartRow <= bRow) {
								final int newtRow = Math.min(tRow + nRow, endRow + 1);
								changeRefIndex(_tRowIndex, tRow, newtRow, ref);
								ref.setTopRow(newtRow);
								if (dstEndRow > bRow) {
									changeRefIndex(_bRowIndex, bRow, dstEndRow, ref);
									ref.setBottomRow(dstEndRow);
								}
							}
						}
					} else { //totally cover
						final int newtRow = tRow + nRow;
						changeRefIndex(_tRowIndex, tRow, newtRow, ref);
						ref.setTopRow(newtRow);
						
						final int newbRow = bRow + nRow;
						changeRefIndex(_bRowIndex, bRow, newbRow, ref);
						ref.setBottomRow(newbRow);
						
						final int newlCol = lCol + nCol;
						changeRefIndex(_lColIndex, lCol, newlCol, ref);
						ref.setLeftCol(newlCol);
						
						final int newrCol = rCol + nCol;
						changeRefIndex(_rColIndex, rCol, newrCol, ref);
						ref.setRightCol(newrCol);
					}
				}
			} else if (tRow >= startRow && bRow <= endRow) { //total cover a column of the Ref
				if (lCol < startCol) {
					if (rCol <= endCol) { //move right side
						if (nRow == 0) {
							if (dstEndCol >= lCol) {
								final int newrCol = Math.max(rCol + nCol, startCol - 1);
								changeRefIndex(_rColIndex, rCol, newrCol, ref);
								ref.setRightCol(newrCol);
								if (dstStartCol < lCol) {
									changeRefIndex(_lColIndex, lCol, dstStartCol, ref);
									ref.setLeftCol(dstStartCol);
								}
							}
						}
					}
				} else {
					if (rCol > endCol) { //move left side
						if (nRow  == 0) {
							if (dstStartCol <= rCol) {
								final int newlCol = Math.min(lCol + nCol, endCol + 1);
								changeRefIndex(_lColIndex, lCol, newlCol, ref);
								ref.setLeftCol(newlCol);
								if (dstEndCol > rCol) {
									changeRefIndex(_rColIndex, rCol, dstEndCol, ref);
									ref.setRightCol(dstEndCol);
								}
							}
						}
					}
				}
			}
		}
		
		//add back the adjusted Refs to ltrb index
		srcHits.addAll(dstHits);
		final Set<Ref> affectHits = new HashSet<Ref>(srcHits);
		affectHits.removeAll(removeHits);
		addOrMergeBackLtrbIndex(affectHits);

		//clear removed refs (Keep the dependents link)
		clearRemoveRefs(removeHits, srcHits);
		
		return getBothDependents(removeHits, srcHits);
	}
	
	@Override
	public void setRefWithIndirectPrecedent(int row, int col, boolean withIndirectPrecedent) {
		final Ref ref0 = getRef(row, col, row, col);
		if (ref0 != null) {
			if (!withIndirectPrecedent) {
				_indirectDependentRefs.remove(ref0);
			} else {
				_indirectDependentRefs.add(ref0);
			}
			ref0.setWithIndirectPrecedent(withIndirectPrecedent);
		}
	}
	
	public String toString(){
		StringBuilder sb = new StringBuilder();
		sb.append("RefSheet:[").append(getOwnerBook().getBookName())
				.append("]").append(getSheetName());
		
		return sb.toString();
	}
	
	@Override
	public Set<Ref> removeExternalRef() {
		Set<Ref> set = new HashSet<Ref>();
		for(Ref ref:new LinkedHashSet<Ref>(_ltrbIndex.values())){//prevent ConcurrentModificationException
			set.addAll(ref.removeExternalRef());
		}
		return set;
	}
}
