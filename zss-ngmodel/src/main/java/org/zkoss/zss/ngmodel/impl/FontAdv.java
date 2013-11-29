package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.NFont;

public abstract class FontAdv implements NFont,Serializable{
	private static final long serialVersionUID = 1L;

	/*package*/ abstract void copyTo(FontAdv dest);

}
