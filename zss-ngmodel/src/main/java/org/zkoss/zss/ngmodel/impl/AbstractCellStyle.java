package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.NCellStyle;

public abstract class AbstractCellStyle implements NCellStyle,Serializable{
	private static final long serialVersionUID = 1L;

	/*package*/ abstract void copyTo(AbstractCellStyle dest);

}
