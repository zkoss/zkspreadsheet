package org.zkoss.zss.ngmodel.impl.sys;

import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

import org.zkoss.zss.ngmodel.impl.RefImpl;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyContext;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyEngine;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.dependency.Ref.RefType;

public class DependencyEngineImpl implements DependencyEngine{

	public DependencyTable getDependencyTable() {
		// TODO Auto-generated method stub
		return new DependencyTable(){

			public Set<Ref> getDependents(Ref precedent) {
				if(precedent.getType()!=RefType.CELL){
					return Collections.EMPTY_SET;
				}
				LinkedHashSet<Ref> dependents = new LinkedHashSet<Ref>();
				//a fake left cell
				RefImpl ref = new RefImpl(precedent.getBookName(),precedent.getSheetName(),precedent.getRow(), precedent.getColumn()+1);
				dependents.add(ref);
				
				return dependents;
			}

			public void clearDependents(Ref dependant) {
				//
			}};
	}

}
