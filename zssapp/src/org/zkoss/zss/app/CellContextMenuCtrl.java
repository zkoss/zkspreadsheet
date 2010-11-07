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
import org.zkoss.zss.app.zul.FontFamily;
import org.zkoss.zss.app.zul.api.Colorbutton;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.Toolbarbutton;
import org.zkoss.zul.Window;
/**
 * @author Sam
 *
 */
public class CellContextMenuCtrl extends GenericForwardComposer{
	FontFamily fontfamily;
	
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

		//Note. Colorbutton use interface, could not auto wired
		_fontColorBtn = (Colorbutton)comp.getFellow("_fontColorBtn");
		_backgroundColorBtn = (Colorbutton)comp.getFellow("_backgroundColorBtn");
	}

	public void onSelect$fontfamily() {
		((Window)spaceOwner).setVisible(false);
	}
	
	public void onChange$_fontColorBtn(Event event) {
		MainWindowCtrl.getInstance().setFontColor(_fontColorBtn.getColor());
		((Window)spaceOwner).setVisible(false);
	}

	public void onChange$_backgroundColorBtn(Event event) {
		MainWindowCtrl.getInstance().setBackgroundColor(_backgroundColorBtn.getColor());
		((Window)spaceOwner).setVisible(false);
	}

	public void onSelect$_fontSizeCombobox(Event event) {
		((Window)spaceOwner).setVisible(false);
		MainWindowCtrl.getInstance().setFontSize(_fontSizeCombobox.getSelectedItem().getLabel());
	}

	public void onBorderSelector(ForwardEvent evt) {
		MainWindowCtrl.getInstance().onBorderSelector(evt);
		((Window)spaceOwner).setVisible(false);
	}
}