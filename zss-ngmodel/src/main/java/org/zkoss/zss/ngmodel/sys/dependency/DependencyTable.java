package org.zkoss.zss.ngmodel.sys.dependency;

import java.util.List;
import java.util.Set;

public interface DependencyTable {

	public Set<Ref> getDependents(Ref precedent);
	
	public void clearDependents(Ref dependant);
}
