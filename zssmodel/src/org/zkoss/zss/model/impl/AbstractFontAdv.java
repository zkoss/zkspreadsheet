/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.model.impl;

import java.io.Serializable;

import org.zkoss.zss.model.SFont;
import org.zkoss.zss.model.SBook;

/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public abstract class AbstractFontAdv implements SFont,Serializable{
	private static final long serialVersionUID = 1L;
	
	/**
	 * gets the string key of this font, the key should combine all the style value in short string as possible
	 */
	abstract String getStyleKey();
	
	//ZSS-1140
	//@since 3.8.2
	abstract public void setOverrideName(boolean b);
	abstract public void setOverrideColor(boolean b);
	abstract public void setOverrideBold(boolean b);
	abstract public void setOverrideItalic(boolean b);
	abstract public void setOverrideStrikeout(boolean b);
	abstract public void setOverrideUnderline(boolean b);
	abstract public void setOverrideHeightPoints(boolean b);
	abstract public void setOverrideTypeOffset(boolean b);
	
    //ZSS-1145
	//@since 3.8.2
	abstract public boolean isOverrideName();
	abstract public boolean isOverrideColor();
	abstract public boolean isOverrideBold();
	abstract public boolean isOverrideItalic();
	abstract public boolean isOverrideStrikeout();
	abstract public boolean isOverrideUnderline();
	abstract public boolean isOverrideHeightPoints();
	abstract public boolean isOverrideTypeOffset();
	
	//ZSS-1183
	//@since 3.9.0
	/*package*/ abstract SFont cloneFont(SBook book);
}
