package org.zkoss.zss.ngmodel.input;

import org.zkoss.zss.ngmodel.NCell.CellType;

public class InputEngine {

	public InputResult parseInput(String input){
		InputResult result = new InputResult();
		if(input!=null){
			result.setType(CellType.STRING);
			result.setValue(input);
		}
		return result;
	}
}
