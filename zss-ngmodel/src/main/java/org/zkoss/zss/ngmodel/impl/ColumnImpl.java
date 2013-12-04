package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NColumn;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.util.CellReference;
import org.zkoss.zss.ngmodel.util.Validations;

public class ColumnImpl extends ColumnAdv {
	private static final long serialVersionUID = 1L;

	private SheetAdv sheet;
	private CellStyleAdv cellStyle;
	
	private Integer width;
	private boolean hidden = false;

	public ColumnImpl(SheetAdv sheet) {
		this.sheet = sheet;
	}

	@Override
	public int getIndex() {
		checkOrphan();
		return sheet.getColumnIndex(this);
	}

	@Override
	public boolean isNull() {
		return false;
	}

	@Override
	void onModelEvent(ModelEvent event) {
		// TODO Auto-generated method stub

	}

	@Override
	public String asString() {
		return CellReference.convertNumToColString(getIndex());
	}

	@Override
	public void checkOrphan() {
		if (sheet == null) {
			throw new IllegalStateException("doesn't connect to parent");
		}
	}

	@Override
	public void release() {
		checkOrphan();
		sheet = null;
	}

	@Override
	public NSheet getSheet() {
		checkOrphan();
		return sheet;
	}

	@Override
	public NCellStyle getCellStyle() {
		return getCellStyle(false);
	}

	@Override
	public NCellStyle getCellStyle(boolean local) {
		if (local || cellStyle != null) {
			return cellStyle;
		}
		checkOrphan();
		return sheet.getBook().getDefaultCellStyle();
	}

	@Override
	public void setCellStyle(NCellStyle cellStyle) {
		Validations.argNotNull(cellStyle);
		Validations.argInstance(cellStyle, CellStyleAdv.class);
		this.cellStyle = (CellStyleAdv) cellStyle;
	}

	@Override
	public int getWidth() {
		if(width!=null){
			return width.intValue();
		}
		checkOrphan();
		return getSheet().getDefaultColumnWidth();
	}

	@Override
	public boolean isHidden() {
		return hidden;
	}

	@Override
	public void setWidth(int width) {
		this.width = Integer.valueOf(width);
	}

	@Override
	public void setHidden(boolean hidden) {
		this.hidden = hidden;
	}

}
