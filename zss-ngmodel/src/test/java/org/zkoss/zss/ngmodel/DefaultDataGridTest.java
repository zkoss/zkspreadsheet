package org.zkoss.zss.ngmodel;

public class DefaultDataGridTest extends ModelTest{
	protected NSheet initialDataGrid(NSheet sheet){
		sheet.setDataGrid(new DefaultDataGrid(sheet));
		return sheet;
	}
}
