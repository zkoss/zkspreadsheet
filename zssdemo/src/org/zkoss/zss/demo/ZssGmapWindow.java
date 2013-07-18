/* ZssGmapWindow.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Dec 6, 2010 2:54:13 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.demo;

import java.text.NumberFormat;
import java.text.ParseException;

import org.zkoss.gmaps.Gmaps;
import org.zkoss.gmaps.Gmarker;
//import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.util.Locales;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.util.GenericForwardComposer;
//import org.zkoss.zss.model.sys.XRanges;
//import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.api.model.CellData.CellType;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.CellEvent;
import org.zkoss.zss.ui.event.EditboxEditingEvent;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.event.StopEditingEvent;
//import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Chart;
import org.zkoss.zul.ChartModel;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.PieModel;
import org.zkoss.zul.SimplePieModel;

/**
 * @author Sam
 *
 */
public class ZssGmapWindow extends GenericForwardComposer {
	Spreadsheet fluSpreadsheet;
	Sheet sheet;
	Gmaps mymap;
	Gmarker[] gmarkerArray;
	Chart myChart;
	final int numOfRows = 42;
	int row, col;
	String prevCellValue;
	NumberFormat format;
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		gmarkerArray = new Gmarker[numOfRows];

		if (fluSpreadsheet == null)
			return;
		fluSpreadsheet.setRowfreeze(1);
		fluSpreadsheet.setColumnfreeze(0);
		
		sheet = fluSpreadsheet.getSelectedSheet();
		
		myChart.setModel(new SimplePieModel());
		updateChart();

		fluSpreadsheet.addEventListener(Events.ON_CELL_FOUCS,
				new EventListener() {
					public void onEvent(Event event) throws Exception {
						onFocus((CellEvent) event);
					}
				});

		fluSpreadsheet.addEventListener(Events.ON_STOP_EDITING,
				new EventListener() {
					public void onEvent(Event event) throws Exception {
						onStopEditingEvent((StopEditingEvent) event);
					}

				});
		fluSpreadsheet.addEventListener(Events.ON_EDITBOX_EDITING,
				new EventListener() {
					public void onEvent(Event event) throws Exception {
						onEditboxEditingEvent((EditboxEditingEvent) event);
					}
				});
//		((SheetImpl) sheet).setSelectionRect(new Rect(0,1,0,1));
//		((SheetImpl) sheet).setFocusRect(new Rect(0,1,0,1));
		format = NumberFormat.getInstance(Locales.getCurrent());
		
		for (int row = 2; row < numOfRows; row++) {
			String state = Ranges.range(sheet, row, 0).getCellEditText();
			// String division = sheet.getCell(row, 1).getCellEditText();

			int numOfCase = (int) format.parse(Ranges.range(sheet, row, 1).getCellEditText()).intValue();

			int numOfDeath = 0;
			try {
				numOfDeath = format.parse(Ranges.range(sheet, row, 2).getCellEditText()).intValue();
			} catch (Exception e) {
			}

			String description = Ranges.range(sheet, row, 3).getCellEditText();
			double lat = format.parse(Ranges.range(sheet, row, 4).getCellEditText()).doubleValue();
			double lng = format.parse(Ranges.range(sheet, row, 5).getCellEditText()).doubleValue();
			String content = "<span style=\"color:#346b93;font-weight:bold\">"
					+ state	+ "</span><br/><span style=\"color:red\">"
					+ numOfCase	+ "</span> cases<br/><span style=\"color:red\">"
					+ numOfDeath + "</span> death<div style=\"background-color:#cfc;padding:2px\">"
					+ description + "</div>";

			Gmarker gmarker = new Gmarker();
			gmarkerArray[row] = gmarker;
			gmarker.setLat(lat);
			gmarker.setLng(lng);
			gmarker.setContent(content);
			mymap.appendChild(gmarker);

			if (row == 2) {
				mymap.setLat(lat);
				mymap.setLng(lng);
				gmarkerArray[row].setOpen(true);
			}
			mymap.setZoom(5);
		}
		
//		self.addEventListener("onEchoInitLater", new EventListener() {
//			public void onEvent(Event event) throws Exception {
//				//Note. In *IE7*: when use gmap, borderlayout, chart and row/column freeze, 
//				// will cause UI error
//				//TODO: cause error
//				fluSpreadsheet.setRowfreeze(1);
//				fluSpreadsheet.setColumnfreeze(0);
//			}
//		});
//		org.zkoss.zk.ui.event.Events.echoEvent(new Event("onEchoInitLater", self, null));
	}
	
	public void onEditboxEditingEvent(EditboxEditingEvent event) throws ParseException {
		if (sheet == null || mymap == null)
			return;

		String str = (String) event.getEditingValue();
		if (col != 1 && col != 2)
			Ranges.range(sheet, row, col).setCellEditText(str);
		if (row != 0)// the header row
			updateRow(row, false);
	}
	
	public void onStopEditingEvent(StopEditingEvent event) throws ParseException {
		if (sheet == null || myChart == null)
			return;
		row = event.getRow();
		col = event.getColumn();
		String str = (String) event.getEditingValue();
		if (col == 1 || col == 2) {
			Double val = null;
			try {
				val = format.parse(str).doubleValue();
				Ranges.range(sheet, row, col).setCellEditText(str);
			} catch (ParseException e) {
				final Integer rowIdx = Integer.valueOf(row);
				final Integer colIdx = Integer.valueOf(col);
				final String prevValue = prevCellValue;
				Messagebox.show("Cell value has to be number format", "Error", 
						Messagebox.OK, Messagebox.ERROR, new EventListener() {
							public void onEvent(Event event) throws Exception {
								Ranges.range(sheet, rowIdx, colIdx).setCellEditText(prevValue);
							}
						});
				return;
			}
		} else {
			Ranges.range(sheet, row, col).setCellEditText(str);
		}
		if (row != 0) {// the header row
			updateRow(row, true);
			if ((col == 1 || col == 2) && (row>=45 && row < 53))
				updateChart();
		}
	}
	
	public void onFocus(CellEvent event) throws ParseException {
		if (mymap == null || sheet == null)
			return;

		Sheet sheet = event.getSheet();
		row = event.getRow();
		col = event.getColumn();
		prevCellValue = Ranges.range(sheet, row, col).getCellEditText();
		if (row < 2 || row > 41)// the header row
			return;

		double lat = format.parse(Ranges.range(sheet, row, 4).getCellEditText()).doubleValue();
		double lng = format.parse(Ranges.range(sheet, row, 5).getCellEditText()).doubleValue();

		mymap.setLat(lat);
		mymap.setLng(lng);
		for (Gmarker gmarker : gmarkerArray) {
			if (gmarker != null && gmarker.isOpen())
				gmarker.setOpen(false);
		}
		gmarkerArray[row].setOpen(true);
	}
	
	public void updateChart(){
		if(myChart == null || sheet == null) return;
		
		ChartModel model = myChart.getModel();
		((PieModel)model).clear();
		for(int row = 45; row < 53; row++){
//			Cell cellName = Utils.getCell(sheet, row, 0);
			String name = Ranges.range(sheet,row,0).getCellData().isBlank()?"":Ranges.range(sheet,row,0).getCellEditText();
//			if(cellName == null || cellName.getStringCellValue() == null)
//				name = "";
//			else
//				name = cellName.getStringCellValue();
			CellData cd = Ranges.range(sheet,row,1).getCellData();
//			Cell cellValue = Utils.getCell(sheet, row, 1);
			Double value;
			if(cd.getResultType()==CellType.NUMERIC){
				value = (Double)cd.getValue();
			}else{
				value = 0D;
			}
			((PieModel)model).setValue(name,value);
			
//			if(cellValue == null)
//				value = new Double(0);
//			else{
//				try{
//					value = (Double)cellValue.getNumericCellValue();
//				}catch(Exception e){
//					value = new Double(0);
//				}
//				((PieModel)model).setValue(name,value);
//			}
		}
	}
	
	public void updateRow(int row, boolean evalValue) throws ParseException {
		if (mymap == null || sheet == null
				|| Ranges.range(sheet, row, 3).getCellData().isBlank())
			return;

		String state = Ranges.range(sheet, row, 0).getCellEditText();
		// String division = sheet.getCell(row, 1).getCellEditText();
		int numOfCase = 0;
		try {
			numOfCase = (int) format.parse(Ranges.range(sheet, row, 1).getCellEditText()).intValue();
		} catch (ParseException e) {
			if (evalValue)
				throw new UiException("Cell value need to be number format");
			else
				return;
		}
		int numOfDeath = 0;
		try {
			numOfDeath = (int) format.parse(Ranges.range(sheet, row, 2).getCellEditText()).intValue();
		} catch (ParseException e) {
			if (evalValue)
				throw new UiException("Cell value need to be number format");
			else
				return;
		}
		String description = Ranges.range(sheet, row, 3).getCellEditText();

		double lat = format.parse(Ranges.range(sheet, row, 4).getCellEditText()).doubleValue();
		double lng = format.parse(Ranges.range(sheet, row, 5).getCellEditText()).doubleValue();
		String content = "<span style=\"color:#346b93;font-weight:bold\">"
				+ state	+ "</span><br/><span style=\"color:red\">"
				+ numOfCase	+ "</span> cases<br/><span style=\"color:red\">"
				+ numOfDeath + "</span> death<div style=\"background-color:#cfc;padding:2px\">"
				+ description + "</div>";

		gmarkerArray[row].setContent(content);
		mymap.setLat(lat);
		mymap.setLng(lng);
		gmarkerArray[row].setOpen(true);
	}
}