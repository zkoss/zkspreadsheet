package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;

public abstract class BookSeriesAdv implements NBookSeries,Serializable{
	private static final long serialVersionUID = 1L;

	/*package*/ abstract DependencyTable getDependencyTable();
}
