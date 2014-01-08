package org.zkoss.zss.ngapi.impl;

import org.zkoss.util.Locales;
import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.format.FormatContext;
import org.zkoss.zss.ngmodel.sys.format.FormatEngine;

public class RangeHelperBase {
	protected final NRange range;
	protected final NSheet sheet;
	private FormatEngine formatEngine;
	
	public RangeHelperBase(NRange range){
		this.range = range;
		this.sheet = range.getSheet();
	}

	public boolean isBlank(NCell cell){
		return cell==null || cell.isNull()||cell.getType() == CellType.BLANK;
	}
	
	public String getFormattedText(NCell cell){
		if(formatEngine==null){
			formatEngine = EngineFactory.getInstance().createFormatEngine();
		}
		
		return formatEngine.format(cell, new FormatContext(Locales.getCurrent())).getText();
	}
	
	public int getRow() {
		return range.getRow();
	}

	public int getColumn() {
		return range.getColumn();
	}

	public int getLastRow() {
		return range.getLastRow();
	}

	public int getLastColumn() {
		return range.getLastColumn();
	}

	public boolean isWholeRow(){
		return range.isWholeRow();
	}
}
