/* ExportComposer.java

	Purpose:
		
	Description:
		
	History:
		November 05, 5:53:16 PM     2010, Created by Ashish Dasnurkar

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
 */
package zss.performance;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.*;
import org.zkoss.zss.api.*;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.CellMouseEvent;

/**
 * @author hawk
 */
public class Fill1MillionComposer extends SelectorComposer {
    @Wire
    Spreadsheet ss;

    @Override
    public void doAfterCompose(Component comp) throws Exception {
        super.doAfterCompose(comp);
        ss.focusTo(1000, 1000);
    }

    @Listen("onCellDoubleClick = #ss")
    public void onCellDoubleClick(CellMouseEvent event) {
        for (int row = 0; row < 1000; row++) {
            for (int column = 0; column < 1000; column++) {
                Range range = Ranges.range(ss.getSelectedSheet(), row, column);
                range.setAutoRefresh(false);
                range.getCellData().setEditText(row + ", " + column);
            }
			Ranges.range(ss.getSelectedSheet(), row, 0, row, 1000).notifyChange();
        }
    }


}
