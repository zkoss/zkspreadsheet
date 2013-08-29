package zss.issue;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.ui.Spreadsheet;

@SuppressWarnings("serial")
public class FreezeComposer extends SelectorComposer<Component>{

	@Wire("spreadsheet")
	private Spreadsheet ss;

	@Listen("onClick = button[label='Freeze Rows']")
	public void freezeRows(){
		ss.setRowfreeze(10);
	}
	
	@Listen("onClick = button[label='Freeze Columns']")
	public void freezeColumns(){
		ss.setColumnfreeze(5);
	}

	@Listen("onClick = button[label='Unfreeze']")
	public void unfreeze(){
		ss.setRowfreeze(-1);
		ss.setColumnfreeze(-1);
	}
	
	@Listen("onClick = button[label='max visible 5 rows']")
	public void setMaxVisibleRow(){
		ss.setMaxrows(5);
	}

	@Listen("onClick = button[label='max visible 25 rows']")
	public void setMaxVisible25Row(){
		ss.setMaxrows(25);
	}
	public void insertFrozenRows(){
		
	}

	public void removeFrozenRows(){

	}

	public void inserFrozenColumns(){

	}

	public void removeFrozenColumns(){

	}
}
