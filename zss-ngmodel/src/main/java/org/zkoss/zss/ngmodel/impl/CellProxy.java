/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel.impl;

import java.lang.ref.WeakReference;

import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NComment;
import org.zkoss.zss.ngmodel.NHyperlink;
import org.zkoss.zss.ngmodel.NRichText;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.util.Validations;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
class CellProxy extends CellAdv {
	private static final long serialVersionUID = 1L;
	private WeakReference<SheetAdv> sheetRef;
	private int rowIdx;
	private int columnIdx;
	CellAdv proxy;

	public CellProxy(SheetAdv sheet, int row, int column) {
		this.sheetRef = new WeakReference<SheetAdv>(sheet);
		this.rowIdx = row;
		this.columnIdx = column;
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

	private void loadProxy() {
		if (proxy == null) {
			proxy = (CellAdv) ((SheetAdv)getSheet()).getCell(rowIdx, columnIdx, false);
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
		return proxy == null ? columnIdx : proxy.getColumnIndex();
	}

	@Override
	public void setValue(Object value) {
		loadProxy();
		if (proxy == null && value != null) {
			proxy = (CellAdv) ((RowAdv) ((SheetAdv)getSheet()).getOrCreateRow(
					rowIdx)).getOrCreateCell(columnIdx);
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
		return proxy == null ? new CellRegion(rowIdx, columnIdx).getReferenceString() : proxy.getReferenceString();
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
		SheetAdv sheet =  ((SheetAdv)getSheet());
		RowAdv row = (RowAdv) sheet.getRow(rowIdx, false);
		NCellStyle style = null;
		if (row != null) {
			style = row.getCellStyle(true);
		}
		if (style == null) {
			ColumnAdv col = (ColumnAdv) sheet.getColumn(columnIdx, false);
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
			proxy = (CellAdv) ((RowAdv)  ((SheetAdv)getSheet()).getOrCreateRow(
					rowIdx)).getOrCreateCell(columnIdx);
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

	@Override
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
	public void destroy() {
		throw new IllegalStateException(
				"never link proxy object and call it's release");
	}

	@Override
	public void checkOrphan() {
	}

	@Override
	public NHyperlink getHyperlink() {
		loadProxy();
		return proxy == null ? null : proxy.getHyperlink();
	}

	@Override
	public void setHyperlink(NHyperlink hyperlink) {
		loadProxy();
		if (proxy == null) {
			proxy = (CellAdv) ((RowAdv)  ((SheetAdv)getSheet()).getOrCreateRow(
					rowIdx)).getOrCreateCell(columnIdx);
		}
		proxy.setHyperlink(hyperlink);
	}
	
	@Override
	public NComment getComment() {
		loadProxy();
		return proxy == null ? null : proxy.getComment();
	}

	@Override
	public void setComment(NComment comment) {
		loadProxy();
		if (proxy == null) {
			proxy = (CellAdv) ((RowAdv)  ((SheetAdv)getSheet()).getOrCreateRow(
					rowIdx)).getOrCreateCell(columnIdx);
		}
		proxy.setComment(comment);
	}

	@Override
	public boolean isFormulaParsingError() {
		loadProxy();
		return proxy == null ? false : proxy.isFormulaParsingError();
	}

	@Override
	public void setRichText(NRichText text) {
		loadProxy();
		if (proxy == null) {
			proxy = (CellAdv) ((RowAdv)  ((SheetAdv)getSheet()).getOrCreateRow(
					rowIdx)).getOrCreateCell(columnIdx);
		}
		proxy.setRichText(text);
	}

	@Override
	public NRichText getRichText() {
		loadProxy();
		return proxy == null ? null : proxy.getRichText();
	}

}
