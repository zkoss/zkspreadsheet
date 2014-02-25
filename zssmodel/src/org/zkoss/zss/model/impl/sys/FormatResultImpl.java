package org.zkoss.zss.model.impl.sys;

import java.text.Format;

import org.zkoss.poi.ss.format.CellFormatResult;
import org.zkoss.zss.model.SColor;
import org.zkoss.zss.model.SRichText;
import org.zkoss.zss.model.impl.ColorImpl;
import org.zkoss.zss.model.sys.format.FormatResult;

public class FormatResultImpl implements FormatResult {
	
	private String text;
	private SColor textColor;//it is possible no format result color
	private SRichText richText;
	private boolean dateFormatted = false;
	private Format formater;
	public FormatResultImpl(SRichText richText){
		this.richText = richText;
	}
	public FormatResultImpl(CellFormatResult result, Format formater, boolean dateFormatted){
		this.text = result.text;
		if (result.textColor != null){
			this.textColor = new ColorImpl((byte)result.textColor.getRed(),(byte)result.textColor.getGreen(),
					(byte)result.textColor.getBlue());
		}
		this.formater = formater;
		this.dateFormatted = dateFormatted;
	}
	public FormatResultImpl(String text, SColor color){
		this.text = text;
		this.textColor = color;
	}
	
	@Override
	public Format getFormater(){
		return formater;
	}
	
	@Override
	public String getText() {
		return richText==null?text:richText.getText();
	}

	@Override
	public SColor getColor() {
		return richText==null?textColor:richText.getFont().getColor();
	}
	@Override
	public boolean isRichText() {
		return richText!=null;
	}
	@Override
	public SRichText getRichText() {
		return richText;
	}
	@Override
	public boolean isDateFormatted() {
		return dateFormatted;
	}


}
