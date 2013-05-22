/* ValidationComposer.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Feb 1, 2012 11:44:19 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
 */
package org.zkoss.zss.app.test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

import org.zkoss.poi.hssf.util.CellRangeAddressList;
import org.zkoss.poi.ss.usermodel.DataValidation;
import org.zkoss.poi.ss.usermodel.DataValidationConstraint;
import org.zkoss.poi.ss.usermodel.DataValidationHelper;
//import org.zkoss.poi.ss.util.CellRangeAddressList;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.util.GenericForwardComposer;
//import org.zkoss.zss.model.sys.XExporter;
//import org.zkoss.zss.model.sys.XExporters;
//import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Filedownload;

/**
 * @author sam
 * 
 */
public class ValidationComposer extends GenericForwardComposer {
	private Spreadsheet spreadsheet;
	private Sheet sheet;

	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		sheet = spreadsheet.getSelectedSheet();
		initValidation();
	}

	private void initValidation() {
		DataValidationHelper dvh = sheet.getPoiSheet().getDataValidationHelper();
		String[] vals = { "STFI", "Treasury", "Fixed Income", "Trade Capture",
				"Unknow" };
		DataValidationConstraint dvc = dvh.createExplicitListConstraint(vals);
		CellRangeAddressList cral = new CellRangeAddressList(1, 100, 3, 3);
		DataValidation dv = dvh.createValidation(dvc, cral);
		sheet.getPoiSheet().addValidationData(dv);
	}

	public void exportExcel() throws IOException {
		Exporter expExcel = Exporters.getExporter("excel");
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		expExcel.export(sheet.getBook(), baos);
		Filedownload.save(baos.toByteArray(), "application/file", sheet
				.getBook().getBookName());
	}
}
