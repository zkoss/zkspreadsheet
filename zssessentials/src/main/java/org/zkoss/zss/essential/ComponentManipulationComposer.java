package org.zkoss.zss.essential;

import org.zkoss.zk.ui.select.annotation.Listen;

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



