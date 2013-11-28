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
	protected AbstractSheet getSheet() {
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
		Validations.argInstance(cellStyle, AbstractCellStyle.class);
		this.cellStyle = (AbstractCellStyle) cellStyle;
	}

}
