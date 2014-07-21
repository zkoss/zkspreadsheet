/* AbstractRefImpl.java

	Purpose:
		
	Description:
		
	History:
		Apr 8, 2010 10:35:05 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.engine.impl;

import java.util.HashSet;
import java.util.Set;

import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.engine.RefSheet;

/**
 * Skeleton implementation for formula references.
 * @author henrichen
 *
 */
public abstract class AbstractRefImpl implements Ref {
	private final RefSheet _ownerSheet;
	private Set<Ref> _dependents;
	private Set<Ref> _precedents;
	
	protected AbstractRefImpl(RefSheet sheet) {
		_ownerSheet = sheet;
	}
	
	@Override
	public RefSheet getOwnerSheet() {
		return _ownerSheet;
	}

	@Override
	public Set<Ref> getPrecedents() {
		if (_precedents == null)
			_precedents = new HashSet<Ref>(3);
		return _precedents; 
	}

	@Override
	public Set<Ref> getDependents() {
		if (_dependents == null)
			_dependents = new HashSet<Ref>(3);
		return _dependents;
	}
	
	@Override
	public void removeAllPrecedents() {
		if (_precedents == null) { //already empty, no need to remove
			return;
		}
		final Set<Ref> precedents = new HashSet<Ref>(_precedents); //avoid comodification
		for(Ref precedent : precedents) {
			this.getPrecedents().remove(precedent);
			precedent.getDependents().remove(this);
		}
		
		//clear orphan references
		for(Ref precedent : precedents) {
			clearIfOrphanRef(precedent);
		}
		clearIfOrphanRef(this);
	}
	
	@Override
	public boolean isWithIndirectPrecedent() {
		return false;
	}
	
	@Override
	public void setWithIndirectPrecedent(boolean b) {
		//ignore
	}
	
	/*package*/ static boolean clearIfOrphanRef(Ref ref) {
		final Set<Ref> dependents = ref.getDependents();
		if (dependents != null && !dependents.isEmpty()) {
			return false; //not empty
		}
		final Set<Ref> precedents = ref.getPrecedents();
		if (precedents != null && !precedents.isEmpty()) {
			return false; //not empty
		}
		//an orphan Ref, remove it from the owner RefSheet
		((AbstractRefImpl)ref).removeSelf();
		return true;
	}
	
	protected abstract void removeSelf();
	
	/*package*/ Ref addPrecedent(String name) {
		final Ref precedent = _ownerSheet.getOwnerBook().getOrCreateVariableRef(name, _ownerSheet);
		this.getPrecedents().add(precedent);
		precedent.getDependents().add(this);
		return precedent;
	}
	
	/*package*/ Ref removePrecedent(String name) {
		final Ref precedent = _ownerSheet.getOwnerBook().getOrCreateVariableRef(name, _ownerSheet);
		if (precedent != null) {
			this.getPrecedents().remove(precedent);
			precedent.getDependents().remove(this);
		}
		return precedent;
	}
	
	/*package*/ Ref addPrecedent(RefSheet refSheet, int tRow, int lCol, int bRow, int rCol) {
		if (refSheet == null) {
			refSheet = _ownerSheet;
		}
		final Ref precedent =  
			refSheet.getOrCreateRef(tRow, lCol, bRow, rCol);
		this.getPrecedents().add(precedent);
		precedent.getDependents().add(this);
		return precedent;
	}
	
	/*package*/ Ref removePrecedent(RefSheet refSheet, int tRow, int lCol, int bRow, int rCol) {
		if (refSheet == null) {
			refSheet = _ownerSheet;
		}
		final AreaRefImpl precedent = (AreaRefImpl) 
			refSheet.getRef(tRow, lCol, bRow, rCol);
		if (precedent != null) {
			this.getPrecedents().remove(precedent);
			precedent.getDependents().remove(this);
		}
		return precedent;
	}
}
