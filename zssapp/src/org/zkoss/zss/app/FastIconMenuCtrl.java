/* FastIconMenuCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Apr 14, 2010 6:43:48 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.app;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.Toolbarbutton;

/**
 * @author Sam
 *
 */
public class FastIconMenuCtrl extends GenericForwardComposer{
	Combobox _fontFamilyCombobox;
	Combobox _fontSizeCombobox;
	Toolbarbutton _boldBtn;
	Toolbarbutton _italicBtn;
	Toolbarbutton _alignLeftBtn;
	Toolbarbutton _alignRightBtn;
	Colorbutton _fontColorBtn;
	Colorbutton _backgroundColorBtn;
	Toolbarbutton _borderBtn;
	Toolbarbutton _mergeCellBtn;
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
	}
	
	public void onChange$_fontColorBtn(Event event) {
		MainWindowCtrl.getInstance().setFontColor(_fontColorBtn.getColor());
	}

	public void onChange$_backgroundColorBtn(Event event) {
		MainWindowCtrl.getInstance().setBackgroundColor(_backgroundColorBtn.getColor());
	}
	
	
}
