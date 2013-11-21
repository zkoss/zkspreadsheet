package org.zkoss.zss.ngmodel.impl;

import java.util.List;

import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NSheet;

public abstract class AbstractBook implements NBook{

	/**
	 * Optimize CellStyle, usually called when export book. 
	 * @return
	 */
	abstract List<NCell> optimizeCellStyle();
}
