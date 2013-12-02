package org.zkoss.zss.ngmodel.sys.format.impl;

import org.zkoss.poi.ss.format.CellFormatResult;
import org.zkoss.zss.ngmodel.sys.format.FormatResult;

public class FormatResultImpl implements FormatResult {
	
	private CellFormatResult formatResult;
	
	public FormatResultImpl(CellFormatResult result){
		formatResult = result;
	}
	
	@Override
	public String getText() {
		return formatResult.text;
	}

	@Override
	public String getColor() {
		return formatResult.textColor.toString();
	}


}
