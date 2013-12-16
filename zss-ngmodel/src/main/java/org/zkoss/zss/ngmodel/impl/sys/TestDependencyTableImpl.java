package org.zkoss.zss.ngmodel.impl.sys;

import java.io.Serializable;
import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.Set;
import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.impl.RefImpl;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.dependency.Ref.RefType;

public class TestDependencyTableImpl extends DependencyTableAdv implements Serializable {
	private static final long serialVersionUID = 1L;

	public Set<Ref> getDependents(Ref precedent) {
		if (precedent.getType() != RefType.CELL) {
			return Collections.EMPTY_SET;
		}
		LinkedHashSet<Ref> dependents = new LinkedHashSet<Ref>();
		// a fake left cell
		RefImpl ref = new RefImpl(precedent.getBookName(),
				precedent.getSheetName(), precedent.getRow(),
				precedent.getColumn() + 1);
		dependents.add(ref);

		return dependents;
	}

	public void clearDependents(Ref dependant) {
		//
	}

	@Override
	public void setBookSeries(NBookSeries series) {
	}

	@Override
	public void add(Ref dependant, Ref precedent) {
	}

	@Override
	public void merge(DependencyTableAdv dependencyTable) {
	}

}
