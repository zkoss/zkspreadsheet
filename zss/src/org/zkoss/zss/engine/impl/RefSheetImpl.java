/* RefSheetImpl.java

	Purpose:
		
	Description:
		
	History:
		Mar 6, 2010 3:39:02 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.engine.impl;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

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
	private final String _sheetName;
	
	public RefSheetImpl(RefBook ownerBook, String sheetName) {
		_ownerBook = ownerBook;
		_sheetName = sheetName;
		_ltrbIndex = new HashMap<RefAddr, Ref>();
		_tRowIndex = new IndexArrayList();
		_bRowIndex = new IndexArrayList();
		_lColIndex = new IndexArrayList();
		_rColIndex = new IndexArrayList();
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
			return Collections.emptySet();
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
			return Collections.emptySet();
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
			return Collections.emptySet();
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
			return Collections.emptySet();
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
				return Collections.emptySet();
		}
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
				_tRowIndex.removeIndexable(tRowRefIndex.getIndex());
			}
		}
		
		final RefIndex bRowRefIndex = (RefIndex) _bRowIndex.getIndexable(bRow);
		if (bRowRefIndex != null) {
			bRowRefIndex.getRefs().remove(ref);
			if (bRowRefIndex.getRefs().isEmpty()) {
				_bRowIndex.removeIndexable(bRowRefIndex.getIndex());
			}
		}
		
		final RefIndex lColRefIndex = (RefIndex) _lColIndex.getIndexable(lCol);
		if (lColRefIndex != null) {
			lColRefIndex.getRefs().remove(ref);
			if (lColRefIndex.getRefs().isEmpty()) {
				_lColIndex.removeIndexable(lColRefIndex.getIndex());
			}
		}
		
		final RefIndex rColRefIndex = (RefIndex) _rColIndex.getIndexable(rCol);
		if (rColRefIndex != null) {
			rColRefIndex.getRefs().remove(ref);
			if (rColRefIndex.getRefs().isEmpty()) {
				_rColIndex.removeIndexable(rColRefIndex.getIndex());
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
		final Set<Ref> tHits = new HashSet<Ref>();
		final Set<Ref> bHits = new HashSet<Ref>();
		
		//bRow hits
		for(int j=_bRowIndex.size() - 1; j >= 0; --j) {
			final RefIndex bRowRefIndex = (RefIndex) _bRowIndex.get(j);  
			final int bRow = bRowRefIndex.getIndex();
			if (bRow < startRow) break; //no more
			bHits.addAll(bRowRefIndex.getRefs());
		}
		hits.addAll(bHits);
		
		//tRow hits
		for(int j=_tRowIndex.size() - 1; j >= 0; --j) {
			final RefIndex tRowRefIndex = (RefIndex) _tRowIndex.get(j);  
			final int tRow = tRowRefIndex.getIndex();
			if (tRow < startRow) break; //no more
			tHits.addAll(tRowRefIndex.getRefs());
		}
		hits.addAll(tHits);

		//remove hits from the ltrbIndex(before adjust row
		for (Ref ref : hits) {
			_ltrbIndex.remove(new RefAddr(ref));
		}

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
				bHits.removeAll(tRowRefs);
				tHits.removeAll(tRowRefs);
			} else {
				tRowRefIndex.setIndex(tRow); //assign a new index
				for(Ref ref : tRowRefs) {
					ref.setTopRow(tRow);
				}
			}
		}

		//add/merge back the adjusted Refs to ltrb index
		final Set<Ref> affectRefs = new HashSet<Ref>(tHits);
		affectRefs.addAll(bHits);
		addOrMergeBackLtrbIndex(affectRefs);

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
		final Set<Ref> bHits = new HashSet<Ref>();
		final Set<Ref> tHits = new HashSet<Ref>();
		
		//bRow hits
		int jb = _bRowIndex.size() - 1;
		for(; jb >= 0; --jb) {
			final RefIndex bRowRefIndex = (RefIndex) _bRowIndex.get(jb);  
			final int bRow = bRowRefIndex.getIndex();
			if (bRow < startRow) break; //no more
			bHits.addAll(bRowRefIndex.getRefs());
		}
		hits.addAll(bHits);
		
		//tRow hits
		int jt = _tRowIndex.size() - 1;
		for(; jt >= 0; --jt) {
			final RefIndex tRowRefIndex = (RefIndex) _tRowIndex.get(jt);  
			final int tRow = tRowRefIndex.getIndex();
			if (tRow < startRow) break; //no more
			tHits.addAll(tRowRefIndex.getRefs());
		}
		hits.addAll(tHits);

		//remove hits from the ltrbIndex(before adjust row)
		for (Ref ref : hits) {
			_ltrbIndex.remove(new RefAddr(ref));
		}
		
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
		for(int j = jt+1, len = _tRowIndex.size(); j < len; ++j) {
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
						bHits.remove(ref);
						tHits.remove(ref);
					} else {
						ref.setTopRow(tRow);
						newRefs.add(ref);
					}
				}
			} else {
				for(Ref ref : tRowRefs) {
					if (ref.getBottomRow() < tRow) { //shall be removed since bRow < tRow
						removeHits.add(ref);
						bHits.remove(ref);
						tHits.remove(ref);
					} else if (min != 0) {
						ref.setTopRow(tRow);
					}
				}
			}
		}
		
		//add back the adjusted Refs to ltrb index
		final Set<Ref> affectHits = new HashSet<Ref>(tHits);
		affectHits.addAll(bHits);
		
		for (Ref ref : affectHits) {
			_ltrbIndex.put(new RefAddr(ref), ref);
		}

		//clear removed refs (Keep the dependents link)
		clearRemoveRefs(removeHits, hits);
		
		return getBothDependents(removeHits, hits);
	}

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
		final Set<Ref> lHits = new HashSet<Ref>();
		final Set<Ref> rHits = new HashSet<Ref>();
		
		//rCol hits
		for(int j=_rColIndex.size() - 1; j >= 0; --j) {
			final RefIndex rColRefIndex = (RefIndex) _rColIndex.get(j);  
			final int rCol = rColRefIndex.getIndex();
			if (rCol < startCol) break; //no more
			rHits.addAll(rColRefIndex.getRefs());
		}
		hits.addAll(rHits);
		
		//lCol hits
		for(int j=_lColIndex.size() - 1; j >= 0; --j) {
			final RefIndex lColRefIndex = (RefIndex) _lColIndex.get(j);  
			final int lCol = lColRefIndex.getIndex();
			if (lCol < startCol) break; //no more
			lHits.addAll(lColRefIndex.getRefs());
		}
		hits.addAll(lHits);

		//remove hits from the ltrbIndex(before adjust column)
		for (Ref ref : hits) {
			_ltrbIndex.remove(new RefAddr(ref));
		}

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
				rHits.removeAll(lColRefs);
				lHits.removeAll(lColRefs);
			} else {
				lColRefIndex.setIndex(lCol); //assign a new index
				for(Ref ref : lColRefs) {
					ref.setLeftCol(lCol);
				}
			}
		}

		//add/merge back the adjusted Refs to ltrb index
		final Set<Ref> affectRefs = new HashSet<Ref>(lHits);
		affectRefs.addAll(rHits);
		addOrMergeBackLtrbIndex(affectRefs);

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
		final Set<Ref> rHits = new HashSet<Ref>();
		final Set<Ref> lHits = new HashSet<Ref>();
		
		//rCol hits
		int jb = _rColIndex.size() - 1;
		for(; jb >= 0; --jb) {
			final RefIndex rColRefIndex = (RefIndex) _rColIndex.get(jb);  
			final int rCol = rColRefIndex.getIndex();
			if (rCol < startCol) break; //no more
			rHits.addAll(rColRefIndex.getRefs());
		}
		hits.addAll(rHits);
		
		//lCol hits
		int jt = _lColIndex.size() - 1;
		for(; jt >= 0; --jt) {
			final RefIndex lColRefIndex = (RefIndex) _lColIndex.get(jt);  
			final int lCol = lColRefIndex.getIndex();
			if (lCol < startCol) break; //no more
			lHits.addAll(lColRefIndex.getRefs());
		}
		hits.addAll(lHits);

		//remove hits from the ltrbIndex(before adjust col)
		for (Ref ref : hits) {
			_ltrbIndex.remove(new RefAddr(ref));
		}
		
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
		for(int j = jt+1, len = _lColIndex.size(); j < len; ++j) {
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
						rHits.remove(ref);
						lHits.remove(ref);
					} else {
						ref.setLeftCol(lCol);
						newRefs.add(ref);
					}
				}
			} else {
				for(Ref ref : lColRefs) {
					if (ref.getRightCol() < lCol) { //shall be removed since rCol < lCol
						removeHits.add(ref);
						rHits.remove(ref);
						lHits.remove(ref);
					} else if (min != 0) {
						ref.setLeftCol(lCol);
					}
				}
			}
		}
		
		//add back the adjusted Refs to ltrb index
		final Set<Ref> affectHits = new HashSet<Ref>(lHits);
		affectHits.addAll(rHits);
		
		for (Ref ref : affectHits) {
			_ltrbIndex.put(new RefAddr(ref), ref);
		}

		//clear removed refs (Keep the dependents link)
		clearRemoveRefs(removeHits, hits);
		
		return getBothDependents(removeHits, hits);
	}

	private void addOrMergeBackLtrbIndex(Set<Ref> affectRefs) {
		for (Ref ref : affectRefs) {
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
	
	private static class RefAddr implements Serializable {
		private static final long serialVersionUID = 201005141203L;
		private int _tRow;
		private int _bRow;
		private int _lCol;
		private int _rCol;
		public RefAddr(int tRow, int lCol, int bRow, int rCol) {
			_tRow = tRow;
			_lCol = lCol;
			_bRow = bRow;
			_rCol = rCol;
		}
		public RefAddr(Ref ref) {
			this(ref.getTopRow(), ref.getLeftCol(), ref.getBottomRow(), ref.getRightCol());
		}
		
		//--Object--//
		@Override
		public int hashCode() {
			return (_lCol == _rCol ? _lCol : (_lCol + _rCol)) 
				+ (_tRow == _bRow ? _tRow : (_tRow + _bRow)) << 14; 
		}
		
		@Override
		public boolean equals(Object other) {
			if (this == other) {
				return true;
			}
			if (!(other instanceof RefAddr)) {
				return false;
			}
			final RefAddr o = (RefAddr) other;
			return _lCol == o._lCol && _rCol == o._rCol && _tRow == o._tRow && _bRow == o._bRow;
		}
	}
}
