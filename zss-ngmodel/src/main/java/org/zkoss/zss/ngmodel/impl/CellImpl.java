package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.util.CellReference;
import org.zkoss.zss.ngmodel.util.Validations;

public class CellImpl extends AbstractCell {

	private AbstractRow row;
	private CellType type = CellType.BLANK;
	private Object value = null;
	private AbstractCellStyle cellStyle;

	public CellImpl(AbstractRow row) {
		this.row = row;
	}

	public CellType getType() {
		return type;
	}

	public boolean isNull() {
		return false;
	}

	public int getRowIndex() {
		checkOrphan();
		return row.getIndex();
	}

	public int getColumnIndex() {
		checkOrphan();
		return row.getCellIndex(this);
	}

	public Object getValue() {
		return value;
	}

	public void setValue(Object value) {
		if (value == null) {
			this.value = value;
			type = CellType.BLANK;
		} else {
			this.value = value.toString();
			type = CellType.STRING;
		}

	}

	public String asString(boolean enableSheetName) {
		return new CellReference(enableSheetName ? row.getSheet()
				.getSheetName() : null, this).formatAsString();
	}

	protected void checkOrphan() {
		if (row == null) {
			throw new IllegalStateException("doesn't connect to parent");
		}
	}

	@Override
	void release() {
		row = null;
	}

	public NCellStyle getCellStyle() {
		return getCellStyle(false);
	}

	public NCellStyle getCellStyle(boolean local) {
		if (local || cellStyle != null) {
			return cellStyle;
		}
		checkOrphan();
		cellStyle = (AbstractCellStyle) row.getCellStyle(true);
		AbstractSheet sheet = row.getSheet();
		if (cellStyle == null) {
			cellStyle = (AbstractCellStyle) sheet.getColumn(getColumnIndex())
					.getCellStyle(true);
		}
		if (cellStyle == null) {
			cellStyle = (AbstractCellStyle) sheet.getBook()
					.getDefaultCellStyle();
		}
		return cellStyle;
	}

	public void setCellStyle(NCellStyle cellStyle) {
		Validations.argNotNull(cellStyle);
		Validations.argInstance(cellStyle, AbstractCellStyle.class);
		this.cellStyle = (AbstractCellStyle) cellStyle;
	}

}
