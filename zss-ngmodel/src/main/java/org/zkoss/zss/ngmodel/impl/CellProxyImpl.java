package org.zkoss.zss.ngmodel.impl;

import java.lang.ref.WeakReference;

import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.util.CellReference;
import org.zkoss.zss.ngmodel.util.Validations;

class CellProxyImpl extends AbstractCell {
	private static final long serialVersionUID = 1L;
	private WeakReference<AbstractSheet> sheetRef;
	private int rowIdx;
	private int columnIdx;
	AbstractCell proxy;

	public CellProxyImpl(AbstractSheet sheet, int row, int column) {
		this.sheetRef = new WeakReference<AbstractSheet>(sheet);
		;
		this.rowIdx = row;
		this.columnIdx = column;
	}

	@Override
	public NSheet getSheet() {
		AbstractSheet sheet = sheetRef.get();
		if (sheet == null) {
			throw new IllegalStateException(
					"proxy target lost, you should't keep this instance");
		}
		return sheet;
	}

	private void loadProxy() {
		if (proxy == null) {
			proxy = (AbstractCell) ((AbstractSheet)getSheet()).getCellAt(rowIdx, columnIdx, false);
			if (proxy != null) {
				sheetRef.clear();
			}
		}
	}

	@Override
	public boolean isNull() {
		loadProxy();
		return proxy == null ? true : proxy.isNull();
	}

	@Override
	public CellType getType() {
		loadProxy();
		return proxy == null ? CellType.BLANK : proxy.getType();
	}

	@Override
	public int getRowIndex() {
		loadProxy();
		return proxy == null ? rowIdx : proxy.getRowIndex();
	}

	@Override
	public int getColumnIndex() {
		loadProxy();
		return proxy == null ? columnIdx : proxy.getRowIndex();
	}

	@Override
	public void setValue(Object value) {
		loadProxy();
		if (proxy == null && value != null) {
			proxy = (AbstractCell) ((AbstractRow) ((AbstractSheet)getSheet()).getOrCreateRowAt(
					rowIdx)).getOrCreateCellAt(columnIdx);
			proxy.setValue(value);
		} else if (proxy != null) {
			proxy.setValue(value);
		}
	}

	@Override
	public Object getValue() {
		loadProxy();
		return proxy == null ? null : proxy.getValue();
	}

	@Override
	public String getReferenceString() {
		loadProxy();
		return proxy == null ? new CellReference(null, rowIdx, columnIdx,
				false, false).formatAsString() : proxy.getReferenceString();
	}

	@Override
	public NCellStyle getCellStyle() {
		return getCellStyle(false);
	}

	@Override
	public NCellStyle getCellStyle(boolean local) {
		loadProxy();
		if (proxy != null) {
			return proxy.getCellStyle(local);
		}
		if (local)
			return null;
		AbstractSheet sheet =  ((AbstractSheet)getSheet());
		AbstractRow row = (AbstractRow) sheet.getRowAt(rowIdx, false);
		NCellStyle style = null;
		if (row != null) {
			style = row.getCellStyle(true);
		}
		if (style == null) {
			AbstractColumn col = (AbstractColumn) sheet.getColumnAt(columnIdx, false);
			if (col != null) {
				style = col.getCellStyle(true);
			}
		}
		if (style == null) {
			style = sheet.getBook().getDefaultCellStyle();
		}
		return style;
	}

	@Override
	public void setCellStyle(NCellStyle cellStyle) {
		Validations.argNotNull(cellStyle);
		loadProxy();
		if (proxy == null) {
			proxy = (AbstractCell) ((AbstractRow)  ((AbstractSheet)getSheet()).getOrCreateRowAt(
					rowIdx)).getOrCreateCellAt(columnIdx);
		}
		proxy.setCellStyle(cellStyle);
	}

	@Override
	public CellType getFormulaResultType() {
		loadProxy();
		return proxy == null ? null : proxy.getFormulaResultType();
	}

	@Override
	public void clearValue() {
		loadProxy();
		if (proxy != null)
			proxy.clearValue();
	}

	public void clearFormulaResultCache() {
		loadProxy();
		if (proxy != null)
			proxy.clearFormulaResultCache();
	}

	@Override
	protected void evalFormula() {
		loadProxy();
		if (proxy != null)
			proxy.evalFormula();
	}

	@Override
	protected Object getValue(boolean eval) {
		loadProxy();
		return proxy == null ? null : proxy.getValue(eval);
	}

	@Override
	public void release() {
		throw new IllegalStateException(
				"never link proxy object and call it's release");
	}

	@Override
	public void checkOrphan() {
	}

}
