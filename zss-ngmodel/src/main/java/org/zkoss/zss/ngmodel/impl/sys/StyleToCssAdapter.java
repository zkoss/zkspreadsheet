package org.zkoss.zss.ngmodel.impl.sys;

import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.sys.format.FormatResult;

/**
 * Convert CellStyle to CSS syntax,
 *  e.g. text-align:right;display: table-cell;vertical-align: middle;font-family:Calibri;color:#000000;
 * @author Hawk
 *
 */
//TODO reference org.zkoss.zss.ui.impl.CellFormatHelper
public class StyleToCssAdapter {
	private NCellStyle cellStyle;
	
	public StyleToCssAdapter(NCellStyle style){
		this.cellStyle = style;
	}

	public String getInnerHtmlSyle(){
		StringBuffer cssResult = new StringBuffer();
		//TODO text weight, alignment 
		
		//TODO get format result from style
		FormatResult formatResult = null;
		
		cssResult.append("color:"+formatResult.getColor().getHtmlColor()+";");
		//TODO format color override font color
		return cssResult.toString();
	}
	
	public String getHtmlStyle(){
		StringBuffer cssResult = new StringBuffer();

		return cssResult.toString();
	}
	
	public boolean hasRightBorder(){
		return false;
	}
	
}
