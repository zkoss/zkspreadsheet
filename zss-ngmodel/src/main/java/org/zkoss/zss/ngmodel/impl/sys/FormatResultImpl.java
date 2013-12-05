package org.zkoss.zss.ngmodel.impl.sys;

import org.zkoss.poi.ss.format.CellFormatResult;
import org.zkoss.zss.ngmodel.NColor;
import org.zkoss.zss.ngmodel.impl.ColorImpl;
import org.zkoss.zss.ngmodel.sys.format.FormatResult;

public class FormatResultImpl implements FormatResult {
	
	private String text;
	private NColor textColor = ColorImpl.BLACK;
	
	public FormatResultImpl(CellFormatResult result){
		this.text = result.text;
		if (result.textColor != null){
			this.textColor = new ColorImpl((byte)result.textColor.getRed(),(byte)result.textColor.getGreen(),
					(byte)result.textColor.getBlue());
		}
	}
	
	@Override
	public String getText() {
		return text;
	}

	@Override
	public NColor getColor() {
		return textColor;
	}


}
