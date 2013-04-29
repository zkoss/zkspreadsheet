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

import org.zkoss.poi.ss.usermodel.DataValidation;
import org.zkoss.poi.ss.usermodel.DataValidationConstraint;
import org.zkoss.poi.ss.usermodel.DataValidationHelper;
import org.zkoss.poi.ss.util.CellRangeAddressList;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.model.sys.Exporter;
import org.zkoss.zss.model.sys.Exporters;
import org.zkoss.zss.model.sys.Worksheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Filedownload;

/**
 * @author sam
 * 
 */
public class ValidationComposer extends GenericForwardComposer {
	private Spreadsheet spreadsheet;
	private Worksheet sheet;

	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		sheet = spreadsheet.getSelectedSheet();
		initValidation();
	}

	private void initValidation() {
		DataValidationHelper dvh = sheet.getDataValidationHelper();
		String[] vals = { "STFI", "Treasury", "Fixed Income", "Trade Capture",
				"Unknow" };
		DataValidationConstraint dvc = dvh.createExplicitListConstraint(vals);
		CellRangeAddressList cral = new CellRangeAddressList(1, 100, 3, 3);
		DataValidation dv = dvh.createValidation(dvc, cral);
		sheet.addValidationData(dv);
	}

	public void exportExcel() {
		Exporter expExcel = Exporters.getExporter("excel");
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		expExcel.export(sheet.getBook(), baos);
		Filedownload.save(baos.toByteArray(), "application/file", sheet
				.getBook().getBookName());
	}
}
