package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NColumn;
import org.zkoss.zss.ngmodel.util.CellReference;
import org.zkoss.zss.ngmodel.util.Validations;

public class ColumnImpl extends AbstractColumn {
	private static final long serialVersionUID = 1L;

	private AbstractSheet sheet;
	private AbstractCellStyle cellStyle;

	public ColumnImpl(AbstractSheet sheet) {
		this.sheet = sheet;
	}

	public int getIndex() {
		checkOrphan();
		return sheet.getColumnIndex(this);
	}

	public boolean isNull() {
		return false;
	}
	@Override
	void onModelEvent(ModelEvent event) {
		// TODO Auto-generated method stub

	}

	public String asString() {
		return CellReference.convertNumToColString(getIndex());
	}

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

	protected AbstractSheet getSheet() {
		checkOrphan();
		return sheet;
	}

	public NCellStyle getCellStyle() {
		return getCellStyle(false);
	}

	public NCellStyle getCellStyle(boolean local) {
		if (local || cellStyle != null) {
			return cellStyle;
		}
		checkOrphan();
		return sheet.getBook().getDefaultCellStyle();
	}

	public void setCellStyle(NCellStyle cellStyle) {
		Validations.argNotNull(cellStyle);
		Validations.argInstance(cellStyle, AbstractCellStyle.class);
		this.cellStyle = (AbstractCellStyle) cellStyle;
	}

}
