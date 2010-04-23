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
import java.util.Map;
import java.util.Set;

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
	private final IndexArrayList _tRowList;
	private final RefBook _ownerBook;
	private final String _sheetName;
	
	public RefSheetImpl(RefBook ownerBook, String sheetName) {
		_ownerBook = ownerBook;
		_tRowList = new IndexArrayList(0); //left top corner row indexes list
		_sheetName = sheetName;
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
		//get or add left top corner column indexes list 
		final IndexArrayList lColList = (IndexArrayList) 
			_tRowList.getOrAddIndexable(new IndexArrayList(tRow));
		
		//get or add right bottom corner row indexes list
		final IndexArrayList bRowList = (IndexArrayList)
			lColList.getOrAddIndexable(new IndexArrayList(lCol));
		
		//get or add right bottom corner column indexes list
		final IndexArrayList rColList = (IndexArrayList)
			bRowList.getOrAddIndexable(new IndexArrayList(bRow));
		
		//get or add a Ref
		final Indexable candidateRef = tRow == bRow && lCol == rCol ?
			new CellRefImpl(tRow, lCol, this) : new AreaRefImpl(tRow, lCol, bRow, rCol, this);
		return (Ref) rColList.getOrAddIndexable(candidateRef); 
	}

	@Override
	public Ref getRef(int tRow, int lCol, int bRow, int rCol) {
		//get left top corner column indexes list 
		final IndexArrayList lColList = (IndexArrayList) 
			_tRowList.getIndexable(tRow);
		
		if (lColList != null) {
			//get right bottom corner row indexes list
			final IndexArrayList bRowList = (IndexArrayList)
				lColList.getIndexable(lCol);
		
			if (bRowList != null) {
				//get right bottom corner column indexes list
				final IndexArrayList rColList = (IndexArrayList)
					bRowList.getIndexable(bRow);

				if (rColList != null) {
					//get the Ref
					return (Ref) rColList.getIndexable(rCol);
				}
			}
		}
		return null;
	}
	
	/* (non-Javadoc)
	 * @see org.zkoss.zss.engine.RefMatrix#getHitRefs(int, int)
	 */
	@Override
	public Set<Ref> getHitRefs(int row, int col) {
		Set<Ref> hits = new HashSet<Ref>();
		for(int j1=0, len1 = _tRowList.size(); j1 < len1; ++j1) {
			final IndexArrayList lColList = (IndexArrayList) _tRowList.get(j1);  
			final int tRow = lColList.getIndex();
			if (tRow > row) break; //no more
			
			for (int j2=0, len2 = lColList.size(); j2 < len2; ++j2) {
				final IndexArrayList bRowList = (IndexArrayList) lColList.get(j2);  
				final int lCol = bRowList.getIndex();
				if (lCol > col) break; //no more
				
				for(int j3= bRowList.size() - 1; j3 >= 0; --j3) {
					final IndexArrayList rColList = (IndexArrayList) bRowList.get(j3);  
					final int bRow = rColList.getIndex();
					if (bRow < row) break; //no more

					for (int j4= rColList.size() - 1; j4 >= 0; --j4) {
						final Indexable ref = (Indexable) rColList.get(j4);  
						final int rCol = ref.getIndex();
						if (rCol < col) break; //no more
						hits.add((Ref)ref); //hit!
					}
				}
			}
		}
		return hits;
	}

	/* (non-Javadoc)
	 * @see org.zkoss.zss.engine.RefMatrix#removeRef(int, int, int, int)
	 */
	@Override
	public Ref removeRef(int tRow, int lCol, int bRow, int rCol) {
		final IndexArrayList lColList = 
			(IndexArrayList) _tRowList.getIndexable(tRow);
		if (lColList != null) {
			final IndexArrayList bRowList = 
				(IndexArrayList) lColList.getIndexable(lCol);
			if (bRowList != null) {
				final IndexArrayList rColList = 
					(IndexArrayList) bRowList.getIndexable(bRow);
				if (rColList != null) {
					return (Ref) rColList.removeIndexable(rCol);
				}
			}
		}
		return null;
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
	
	/**
	 * Index array used to manage object with indexes (such as formula references).
	 * @author henrichen
	 */
	private static class IndexArrayList extends ArrayList<Indexable> implements Indexable  {
		private static final long serialVersionUID = 201003061603L;
		private int _index;
		
		public IndexArrayList(int index) {
			_index = index;
		}
		
		public int getIndex() {
			return _index;
		}
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
		 * Remove the indexable in the speicifed index and return it; return null if not exist.
		 * @param index the indexable index
 		 * @return the removed indexable
		 */
		public Indexable removeIndexable(int index) {
			final int j = Collections.binarySearch(this, new DummyIndexable(index));
			return j < 0 ? null : remove(j); 
		}

		@Override
		public int compareTo(Indexable o) {
			return this.getIndex() - o.getIndex();
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
}
