/* DependencyTrackerHelper.java

	Purpose:
		
	Description:
		
	History:
		Apr 8, 2010 9:36:45 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.engine.impl;

import java.util.HashSet;
import java.util.Set;

import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.engine.RefSheet;

/**
 * Helper class for dependency tracking and operation.
 * @author henrichen
 *
 */
public class DependencyTrackerHelper {
	public static void getBothDependents(RefSheet sheet, int row, int col, Set<Ref> all, Set<Ref> last) {
		final Set<Ref> dependents = getDirectDependents(sheet, row, col);
		getBothDependents(dependents, all, last);
	}
	
	public static void getBothDependents(Set<Ref> dependents, Set<Ref> all, Set<Ref> last) {
		for (Ref dependent : dependents) {
			getDependentsRecursive(dependent, all, last); //recursive
		}
	}
	
	public static Set<Ref> getDirectDependents(RefSheet sheet, int row, int col) {
		final Set<Ref> refs = sheet.getHitRefs(row, col);
		final Set<Ref> dependents = new HashSet<Ref>();
		for(Ref ref: refs) {
			dependents.addAll(ref.getDependents());
		}
		return dependents;
	}
	
	//TODO can form a cell link per the formula
	private static void getDependentsRecursive(Ref srcRef, Set<Ref> all, Set<Ref> last) {
		if (all.contains(srcRef)) return; //circular graph, return to avoid endless loop
		all.add(srcRef);
		final Set<Ref> dependents = getDirectDependents(srcRef.getOwnerSheet(), srcRef.getTopRow(), srcRef.getLeftCol());
		if (last != null && dependents.isEmpty()) { //no more dependents, so I am the last
			last.add(srcRef);
		}
		for (Ref dependent : dependents) {
			getDependentsRecursive(dependent, all, last); //recursive
		}
	}

	public static void addDependency(CellRefImpl srcRef, RefSheet sheet,
		int tRow, int lCol, int bRow, int rCol) {
		if (sheet == null) {
			sheet = srcRef.getOwnerSheet();
		}
		srcRef.addPrecedent(sheet, tRow, lCol, bRow, rCol);
	}

	public static void removeDependency(CellRefImpl srcRef, RefSheet sheet, 
		int tRow, int lCol, int bRow, int rCol) {
		if (srcRef != null) {
			if (sheet == null) {
				sheet = srcRef.getOwnerSheet();
			}
			final Ref precedent = srcRef.removePrecedent(sheet, tRow, lCol, bRow, rCol);
			if (precedent.getDependents().isEmpty() && precedent.getPrecedents().isEmpty()) {
				sheet.removeRef(tRow, lCol, bRow, rCol);
			}
			if (srcRef.getDependents().isEmpty() && srcRef.getPrecedents().isEmpty()) {
				removeRef(srcRef);
			}
		}
	}
	
	public static void addDependency(CellRefImpl srcRef, String name) {
		srcRef.addPrecedent(name);
	}

	public static void removeDependency(CellRefImpl srcRef, String name) {
		if (srcRef != null) {
			final Ref precedent = srcRef.removePrecedent(name);
			if (precedent.getDependents().isEmpty() && precedent.getPrecedents().isEmpty()) {
				srcRef.getOwnerSheet().getOwnerBook().removeVariableRef(name);
			}
			if (srcRef.getDependents().isEmpty() && srcRef.getPrecedents().isEmpty()) {
				removeRef(srcRef);
			}
		}
	}
	
	private static void removeRef(CellRefImpl srcRef) {
		final int srcRow = srcRef.getTopRow();
		final int srcCol = srcRef.getLeftCol();
		srcRef.getOwnerSheet().removeRef(srcRow, srcCol, srcRow, srcCol);
	}
}
