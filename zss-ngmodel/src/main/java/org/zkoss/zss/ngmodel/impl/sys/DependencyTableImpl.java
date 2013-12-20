package org.zkoss.zss.ngmodel.impl.sys;

import java.util.EnumSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.LinkedList;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Queue;
import java.util.Set;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.dependency.Ref.RefType;

/* DependencyTableImpl.java

 Purpose:

 Description:

 History:
 Nov 22, 2013 Created by Pao Wang

 Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */

/**
 * Default implementation of dependency table.
 * @author Pao
 */
public class DependencyTableImpl extends DependencyTableAdv {
	private static final long serialVersionUID = 1L;
	private static final EnumSet<RefType> regionTypes = EnumSet.of(RefType.BOOK, RefType.SHEET, RefType.AREA,
			RefType.CELL);

	/** Map<dependant, precedent> */
	private Map<Ref, Set<Ref>> map = new LinkedHashMap<Ref, Set<Ref>>();
	private NBookSeries books;

	public DependencyTableImpl() {
	}

	@Override
	public void setBookSeries(NBookSeries series) {
		this.books = series;
	}

	@Override
	public void add(Ref dependant, Ref precedent) {
		Set<Ref> precedents = map.get(dependant);
		if(precedents == null) {
			precedents = new LinkedHashSet<Ref>();
			map.put(dependant, precedents);
		}
		precedents.add(precedent);
	}

	public void clear() {
		map.clear();
	}

	@Override
	public void clearDependents(Ref dependant) {
		if(dependant != null) {
			map.remove(dependant);
		}
	}

	@Override
	public Set<Ref> getDependents(Ref precedent) {

		// search dependents and their dependents recursively
		Set<Ref> result = new LinkedHashSet<Ref>();
		Queue<Ref> queue = new LinkedList<Ref>();
		queue.add(precedent);

		while(!queue.isEmpty()) {
			Ref p = queue.remove();
			for(Entry<Ref, Set<Ref>> entry : map.entrySet()) {
				Ref target = entry.getKey();
				if(!result.contains(target)) {
					for(Ref pre : entry.getValue()) {
						if(isMatched(pre, p)) {
							result.add(target);
							queue.add(target);
							break;
						}
					}
				}
			}
		}
		return result;
	}

	private boolean isMatched(Ref a, Ref b) {
		if(regionTypes.contains(a.getType()) && regionTypes.contains(b.getType())) {
			return isIntersected(a, b);
		} else {
			return a.equals(b);
		}
	}

	/**
	 * @return true if b overlaps a.
	 */
	private boolean isIntersected(Ref a, Ref b) {

		// check book is the same or not
		if(!a.getBookName().equals(b.getBookName())) {
			return false;
		}

		// anyone is a book, matched immediately
		if(a.getType() == RefType.BOOK || b.getType() == RefType.BOOK) {
			return true;
		}

		// check sheets are intersected or not
		// just assume 3D reference
		NBook book = books.getBook(a.getBookName());
		int[] aSheetIndexes = getSheetIndex(book, a);
		int[] bSheetIndexes = getSheetIndex(book, b);
		if(!isIntersected(aSheetIndexes[0], aSheetIndexes[1], bSheetIndexes[0], bSheetIndexes[1])) {
			return false;
		}

		// anyone is a sheet, matched immediately
		if(a.getType() == RefType.SHEET || b.getType() == RefType.SHEET) {
			return true;
		}

		// Okay, they only can be area or cell now!
		// check overlapped or not
		return isIntersected(a.getColumn(), a.getRow(), a.getLastColumn(), a.getLastRow(), b.getColumn(),
				b.getRow(), b.getLastColumn(), b.getLastRow());
	}

	private int[] getSheetIndex(NBook book, Ref ref) {
		NSheet sheet = book.getSheetByName(ref.getSheetName());
		int a = sheet != null ? book.getSheetIndex(sheet) : -1;
		sheet = book.getSheetByName(ref.getLastSheetName());
		int b = sheet != null ? book.getSheetIndex(sheet) : a;
		return new int[]{a, b}; // Excel always adjust 3D formula to ascending, we assume this too.
	}

	/** point in line */
	private final boolean isIntersected(int a1, int a2, int b) {
		return (a1 <= b && b <= a2);
	}

	/** line overlaps line */
	private final boolean isIntersected(int a1, int a2, int b1, int b2) {
		return isIntersected(a1, a2, b1) || isIntersected(a1, a2, b2) || isIntersected(b1, b2, a1) || isIntersected(
				b1, b2, a2);
	}

	/** rect. overlaps rect. */
	private final boolean isIntersected(int ax1, int ay1, int ax2, int ay2, int bx1, int by1, int bx2, int by2) {
		return (isIntersected(ax1, ax2, bx1, bx2) && isIntersected(ay1, ay2, by1, by2));
	}

	@Override
	public String toString() {
		StringBuilder sb = new StringBuilder();
		for(Entry<Ref, Set<Ref>> entry : map.entrySet()) {
			Ref target = entry.getKey();
			sb.append(target).append('\n');
			for(Ref pre : entry.getValue()) {
				sb.append('\t').append(pre).append('\n');
			}
		}
		return sb.toString();
	}

	@Override
	public void merge(DependencyTableAdv dependencyTable) {
		// TODO Auto-generated method stub
	}
}
