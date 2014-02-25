package org.zkoss.zss.ngapi.impl;

import org.zkoss.util.Locales;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SCell.CellType;
import org.zkoss.zss.model.sys.EngineFactory;
import org.zkoss.zss.model.sys.format.FormatContext;
import org.zkoss.zss.model.sys.format.FormatEngine;
import org.zkoss.zss.model.sys.formula.FormulaEngine;
import org.zkoss.zss.ngapi.NRange;

public class RangeHelperBase {
	protected final NRange range;
	protected final SSheet sheet;
	private FormatEngine formatEngine;
	private FormulaEngine formulaEngine;
	
	public RangeHelperBase(NRange range){
		this.range = range;
		this.sheet = range.getSheet();
	}

	public static boolean isBlank(SCell cell){
		return cell==null || cell.isNull()||cell.getType() == CellType.BLANK;
	}
	
	protected FormatEngine getFormatEngine(){
		if(formatEngine==null){
			formatEngine = EngineFactory.getInstance().createFormatEngine();
		}
		return formatEngine;
	}
	
	public String getFormattedText(SCell cell){
		return getFormatEngine().format(cell, new FormatContext(Locales.getCurrent())).getText();
	}
	
	protected FormulaEngine getFormulaEngine(){
		if (formulaEngine == null){
			formulaEngine = EngineFactory.getInstance().createFormulaEngine();
		}
		return formulaEngine;
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
	
	public boolean isWholeColumn() {
		return range.isWholeColumn();
	}
}
