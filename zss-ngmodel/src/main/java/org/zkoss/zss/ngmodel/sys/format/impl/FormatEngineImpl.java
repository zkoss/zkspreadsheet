package org.zkoss.zss.ngmodel.sys.format.impl;

import org.zkoss.poi.ss.format.CellFormat;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.sys.format.FormatContext;
import org.zkoss.zss.ngmodel.sys.format.FormatEngine;
import org.zkoss.zss.ngmodel.sys.format.FormatResult;

public class FormatEngineImpl implements FormatEngine {

	@Override
	public FormatResult format(NCell cell, FormatContext context) {
		CellFormat formatter = CellFormat.getInstance(cell.getCellStyle().getDataFormat(), context.getLocale());
		return new FormatResultImpl(formatter.apply(cell.getValue()));
	}

}
