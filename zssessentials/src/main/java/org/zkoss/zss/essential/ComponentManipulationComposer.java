package org.zkoss.zss.essential;

import java.util.Arrays;
import java.util.List;
import org.zkoss.zel.impl.util.Objects;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.RangeRunner;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.EditableCellStyle;
import org.zkoss.zss.essential.util.ClientUtil;
import org.zkoss.zss.ui.Position;
import org.zkoss.zul.Label;
import org.zkoss.zul.Textbox;

/**
 * This class controll the spreadsheet component
 * @author dennis
 *
 */
public class ComponentManipulationComposer extends AbstractDemoComposer {
	
	@Listen("onClick = #toggleFreeze")
	public void onToggleFreeze(){
		if(ss.getRowfreeze()>=0){
			ss.setRowfreeze(-1);
			ss.setColumnfreeze(-1);
		}else{
			ss.setRowfreeze(3);
			ss.setColumnfreeze(2);
		}
	}
	
}



