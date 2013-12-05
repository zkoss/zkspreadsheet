package org.zkoss.zss.ngmodel.impl;

import java.lang.ref.WeakReference;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.util.CellReference;
import org.zkoss.zss.ngmodel.util.Validations;

class ColumnProxy extends ColumnAdv {
	private static final long serialVersionUID = 1L;
	private final WeakReference<SheetAdv> sheetRef;
	private final int index;
	private ColumnAdv proxy;

	public ColumnProxy(SheetAdv sheet, int index) {
		this.sheetRef = new WeakReference(sheet);
		this.index = index;
	}

	protected void loadProxy() {
		if (proxy == null) {
			proxy = (ColumnAdv) ((SheetAdv)getSheet()).getColumn(index, false);
			if (proxy != null) {
				sheetRef.clear();
			}
		}
	}

	@Override
	public NSheet getSheet() {
		SheetAdv sheet = sheetRef.get();
		if (sheet == null) {
			throw new IllegalStateException(
					"proxy target lost, you should't keep this instance");
		}
		return sheet;
	}

	@Override
	public int getIndex() {
		loadProxy();
		return proxy == null ? index : proxy.getIndex();
	}

	@Override
	public boolean isNull() {
		loadProxy();
		return proxy == null ? true : proxy.isNull();
	}

	@Override
	public String asString() {
		loadProxy();
		return proxy == null ? CellReference.convertNumToColString(getIndex())
				: proxy.asString();
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
		return local ? null : getSheet().getBook().getDefaultCellStyle();
	}

	@Override
	public void setCellStyle(NCellStyle cellStyle) {
		Validations.argNotNull(cellStyle);
		loadProxy();
		if (proxy == null) {
			proxy = (ColumnAdv)((SheetAdv)getSheet()).getOrCreateColumn(index);
		}
		proxy.setCellStyle(cellStyle);
	}

	@Override
	public void destroy() {
		throw new IllegalStateException(
				"never link proxy object and call it's release");
	}

	@Override
	public void checkOrphan() {}

	@Override
	void onModelEvent(ModelEvent event) {}

	@Override
	public int getWidth() {
		loadProxy();
		if (proxy != null) {
			return proxy.getWidth();
		}
		return getSheet().getDefaultColumnWidth();
	}

	@Override
	public boolean isHidden() {
		loadProxy();
		if (proxy != null) {
			return proxy.isHidden();
		}
		return false;
	}

	@Override
	public void setWidth(int width) {
		loadProxy();
		if (proxy == null) {
			proxy = (ColumnAdv)((SheetAdv)getSheet()).getOrCreateColumn(index);
		}
		proxy.setWidth(width);
	}

	@Override
	public void setHidden(boolean hidden) {
		loadProxy();
		if (proxy == null) {
			proxy = (ColumnAdv)((SheetAdv)getSheet()).getOrCreateColumn(index);
		}
		proxy.setHidden(hidden);
	}
}
