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
package org.zkoss.zss.ngapi.impl.imexp;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.zkoss.zss.ngapi.ImporterFactory;
import org.zkoss.zss.ngapi.NImporter;
import org.zkoss.zss.ngapi.NRanges;
import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.DefaultDataGrid;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBooks;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NCellStyle.FillPattern;
import org.zkoss.zss.ngmodel.NCellStyle.VerticalAlignment;
import org.zkoss.zss.ngmodel.NCellValue;
import org.zkoss.zss.ngmodel.NChart;
import org.zkoss.zss.ngmodel.NChart.NChartLegendPosition;
import org.zkoss.zss.ngmodel.NChart.NChartType;
import org.zkoss.zss.ngmodel.NPicture.Format;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.NCellStyle.Alignment;
import org.zkoss.zss.ngmodel.NViewAnchor;
import org.zkoss.zss.ngmodel.chart.NGeneralChartData;
import org.zkoss.zss.ngmodel.chart.NSeries;
import org.zkoss.zss.ngmodel.impl.BookImpl;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class TestImporterFactory implements ImporterFactory{

	static NBook book;//test test share
	
	@Override
	public NImporter createImporter() {
		return new AbstractImporter() {
			
			@Override
			public NBook imports(InputStream is, String bookName) throws IOException {
//				if(book!=null){
//					return book;
//				}
//				book = NBooks.createBook(bookName);
//				book.setShareScope("application");
				
				NBook book = NBooks.createBook(bookName);
				
				buildDataGridSheet(book);
				
				buildChartSheet(book);
				
				buildNormalSheet(book);

				buildFreeze(book);
				return book;
			}

			private void buildDataGridSheet(NBook book) {
				NSheet sheet = book.createSheet("DataGrid");
				sheet.setDataGrid(new DefaultDataGrid(sheet));
				sheet.getDataGrid().setValue(0, 0, new NCellValue(CellType.BLANK,null));
				sheet.getDataGrid().setValue(0, 0, new NCellValue(CellType.STRING,"ABC"));
				sheet.getDataGrid().setValue(1, 1, new NCellValue(CellType.NUMBER,12));
				sheet.getDataGrid().setValue(2, 2, new NCellValue(CellType.NUMBER,3.45));
				sheet.getDataGrid().setValue(2, 2, new NCellValue(CellType.BOOLEAN,false));
			}

			private void buildFreeze(NBook book) {
				NSheet sheet = book.createSheet("Freeze");
				
				sheet.getViewInfo().setNumOfColumnFreeze(5);
				sheet.getViewInfo().setNumOfRowFreeze(7);
				sheet.addPicture(Format.JPG, getTestImageData(), new NViewAnchor(3, 3, 600, 300));
			}

			private void buildChartSheet(NBook book) {
				NSheet sheet = book.createSheet("Chart");
				
				sheet.getViewInfo().setNumOfRowFreeze(6);
				
				sheet.getCell(0, 0).setValue("My Series");
				sheet.getCell(0, 1).setValue("Volumn");
				sheet.getCell(0, 2).setValue("Open");
				sheet.getCell(0, 3).setValue("High");
				sheet.getCell(0, 4).setValue("Low");
				sheet.getCell(0, 5).setValue("Close");
				
				
				sheet.getCell(1, 0).setValue("A");
				sheet.getCell(2, 0).setValue("B");
				sheet.getCell(3, 0).setValue("C");
				sheet.getCell(4, 0).setValue("D");
				sheet.getCell(5, 0).setValue("E");
				sheet.getCell(6, 0).setValue("F");
				
				NCellStyle style = book.createCellStyle(true);
				style.setDataFormat("yyyy/m/d");
				sheet.getCell(1, 7).setValue(newDate("2013/1/1"));
				sheet.getCell(1, 7).setCellStyle(style);
				sheet.getCell(2, 7).setValue(newDate("2013/1/2"));
				sheet.getCell(2, 7).setCellStyle(style);
				sheet.getCell(3, 7).setValue(newDate("2013/1/3"));
				sheet.getCell(3, 7).setCellStyle(style);
				sheet.getCell(4, 7).setValue(newDate("2013/1/4"));
				sheet.getCell(4, 7).setCellStyle(style);
				sheet.getCell(5, 7).setValue(newDate("2013/1/5"));
				sheet.getCell(5, 7).setCellStyle(style);
				sheet.getCell(6, 7).setValue(newDate("2013/1/6"));
				sheet.getCell(6, 7).setCellStyle(style);
				
				sheet.getCell(1, 1).setValue(1);
				sheet.getCell(2, 1).setValue(2);
				sheet.getCell(3, 1).setValue(3);
				sheet.getCell(4, 1).setValue(1);
				sheet.getCell(5, 1).setValue(2);
				sheet.getCell(6, 1).setValue(3);
				
				sheet.getCell(1, 2).setValue(4);
				sheet.getCell(2, 2).setValue(5);
				sheet.getCell(3, 2).setValue(6);
				sheet.getCell(4, 2).setValue(1);
				sheet.getCell(5, 2).setValue(2);
				sheet.getCell(6, 2).setValue(3);
				
				sheet.getCell(1, 3).setValue(7);
				sheet.getCell(2, 3).setValue(8);
				sheet.getCell(3, 3).setValue(9);
				sheet.getCell(4, 3).setValue(2);
				sheet.getCell(5, 3).setValue(2);
				sheet.getCell(6, 3).setValue(3);
				
				sheet.getCell(1, 4).setValue(1);
				sheet.getCell(2, 4).setValue(3);
				sheet.getCell(3, 4).setValue(5);
				sheet.getCell(4, 4).setValue(2);
				sheet.getCell(5, 4).setValue(2);
				sheet.getCell(6, 4).setValue(3);
				
				sheet.getCell(1, 5).setValue(2);
				sheet.getCell(2, 5).setValue(6);
				sheet.getCell(3, 5).setValue(9);
				sheet.getCell(4, 5).setValue(3);
				sheet.getCell(5, 5).setValue(2);
				sheet.getCell(6, 5).setValue(3);
				
				sheet.getCell(1, 6).setValue(1);
				sheet.getCell(2, 6).setValue(4);
				sheet.getCell(3, 6).setValue(8);
				sheet.getCell(4, 6).setValue(3);
				sheet.getCell(5, 6).setValue(2);
				sheet.getCell(6, 6).setValue(3);
				
				
				NChart chart = sheet.addChart(NChartType.PIE, new NViewAnchor(1, 12, 300, 200));
				buildChartData(chart);
				chart.setLegendPosition(NChartLegendPosition.RIGHT);
				
				chart = sheet.addChart(NChartType.PIE, new NViewAnchor(1, 18, 300, 200));
				buildChartData(chart);
				chart.setTitle("Another Title");
				chart.setThreeD(true);
				
				chart = sheet.addChart(NChartType.BAR, new NViewAnchor(12, 0, 300, 200));
				buildChartData(chart);
				chart.setLegendPosition(NChartLegendPosition.BOTTOM);
				
				chart = sheet.addChart(NChartType.BAR, new NViewAnchor(12, 6, 300, 200));
				buildChartData(chart);
				chart.setThreeD(true);
				
				chart = sheet.addChart(NChartType.COLUMN, new NViewAnchor(12, 12, 300, 200));
				buildChartData(chart);
				chart.setLegendPosition(NChartLegendPosition.LEFT);
				
				chart = sheet.addChart(NChartType.COLUMN, new NViewAnchor(12, 18, 300, 200));
				buildChartData(chart);
				chart.setThreeD(true);
				
				chart = sheet.addChart(NChartType.DOUGHNUT, new NViewAnchor(12, 18, 300, 200));
				buildChartData(chart);
				chart.setLegendPosition(NChartLegendPosition.RIGHT);
				
				
				chart = sheet.addChart(NChartType.LINE, new NViewAnchor(24, 0, 300, 200));
				buildChartData(chart);
				chart.setLegendPosition(NChartLegendPosition.TOP);
				
				chart = sheet.addChart(NChartType.LINE, new NViewAnchor(24, 6, 300, 200));
				buildChartData(chart);
				chart.setThreeD(true);
				
				chart = sheet.addChart(NChartType.AREA, new NViewAnchor(24, 12, 300, 200));
				buildChartData(chart);
				
				chart = sheet.addChart(NChartType.SCATTER, new NViewAnchor(36, 0, 300, 200));
				buildScatterChartData(chart);
				chart.setLegendPosition(NChartLegendPosition.RIGHT);
				
				chart = sheet.addChart(NChartType.BUBBLE, new NViewAnchor(36, 6, 300, 200));
				buildBubbleChartData(chart);
				chart.setLegendPosition(NChartLegendPosition.RIGHT);
//				
				chart = sheet.addChart(NChartType.STOCK, new NViewAnchor(36, 12, 600, 200));
				buildStockChartData(chart);
				
				
			}
			
			private Date newDate(String string) {
				SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
				try {
					return sdf.parse(string);
				} catch (ParseException e) {
					e.printStackTrace();
				}
				return null;
			}

			private void buildStockChartData(NChart chart){//X,Y
				NGeneralChartData data = (NGeneralChartData)chart.getData();
				data.setCategoriesFormula("H2:H7");
				NSeries series = data.addSeries();//volumn
				series.setFormula("B1", "B2:B7");
				series = data.addSeries();
				series.setFormula("C1", "C2:C7");//open
				series = data.addSeries();
				series.setFormula("D1", "D2:D7");//high
				series = data.addSeries();
				series.setFormula("E1", "E2:E7");//low
				series = data.addSeries();
				series.setFormula("F1", "F2:F7");//close
			}
			
			private void buildBubbleChartData(NChart chart){//X,Y
				NGeneralChartData data = (NGeneralChartData)chart.getData();
				NSeries series = data.addSeries();
				series.setXYZFormula("A3", "B2:G2", "B3:G3", "B5:G5");
				series = data.addSeries();
				series.setXYZFormula("A4", "B2:G2", "B4:G4", "B5:G5");
			}
			
			private void buildScatterChartData(NChart chart){//X,Y
				NGeneralChartData data = (NGeneralChartData)chart.getData();
				NSeries series = data.addSeries();
				series.setXYFormula("A3", "B2:G2", "B3:G3");
				series = data.addSeries();
				series.setXYFormula("A4", "B2:G2", "B4:G4");
			}			
			
			private void buildChartData(NChart chart){
				NGeneralChartData data = (NGeneralChartData)chart.getData();
				data.setCategoriesFormula("A2:A4");//A,B,C
				NSeries series = data.addSeries();
				series.setXYFormula("A1", "B2:B4", null);
				series = data.addSeries();
				series.setXYFormula("\"Series 2\"", "C2:C4", null);
				series = data.addSeries();
				series.setXYFormula(null, "D2:D4", null);
			}

			private void buildNormalSheet(NBook book) {
				NSheet sheet = book.createSheet("Sheet 1");
				sheet.getColumn(0).setWidth(120);
				sheet.getColumn(1).setWidth(120);
				sheet.getColumn(2).setWidth(120);
				sheet.getColumn(3).setWidth(120);
				sheet.getColumn(4).setWidth(120);
				sheet.getColumn(5).setWidth(120);
				sheet.getColumn(6).setWidth(120);
				
				sheet.getCell(0,11).setStringValue("Column M,O is hidden");
				sheet.getColumn(12).setHidden(true);
				sheet.getColumn(14).setHidden(true);
				
				sheet.getCell(15,0).setStringValue("Row 17,19 is hidden");
				sheet.getRow(16).setHidden(true);
				sheet.getRow(18).setHidden(true);
				
				NCellStyle style;
				NCell cell;
				SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");
				Date now = new Date();
				Date dayonly = null;
				try {
					dayonly = sdf.parse(sdf.format(now));
				} catch (ParseException e) {
					e.printStackTrace();
				}
				
				(cell = sheet.getCell(0, 0)).setValue("Values:");
				(cell = sheet.getCell(0, 1)).setValue(123.45);
				(cell = sheet.getCell(0, 2)).setValue(now);
				(cell = sheet.getCell(0, 3)).setValue(Boolean.TRUE);
				
				(cell = sheet.getCell(1, 0)).setValue("Number Format:");
				(cell = sheet.getCell(1, 1)).setValue(33);
				style = book.createCellStyle(true);
				style.setDataFormat("0.00");
				cell.setCellStyle(style);

				(cell = sheet.getCell(1, 2)).setValue(44.55);
				style = book.createCellStyle(true);
				style.setDataFormat("$#,##0.0");
				cell.setCellStyle(style);
				
				
				(cell = sheet.getCell(1, 3)).setValue(77.88);
				style = book.createCellStyle(true);
				style.setDataFormat("0.00;[Red]0.00");
				cell.setCellStyle(style);
				
				(cell = sheet.getCell(1, 4)).setValue(-77.88);
				style = book.createCellStyle(true);
				style.setDataFormat("0.00;[Red]0.00");
				cell.setCellStyle(style);
				
				
				(cell = sheet.getCell(2, 0)).setValue("Date Format:");
				(cell = sheet.getCell(2, 1)).setValue(dayonly);
				style = book.createCellStyle(true);
				style.setDataFormat("yyyy/m/d");
				cell.setCellStyle(style);
				
				(cell = sheet.getCell(2, 2)).setValue(dayonly);
				style = book.createCellStyle(true);
				style.setDataFormat("m/d/yyy");
				cell.setCellStyle(style);
				
				(cell = sheet.getCell(2, 3)).setValue(now);
				style = book.createCellStyle(true);
				style.setDataFormat("m/d/yy h:mm;@");
				cell.setCellStyle(style);
				
				(cell = sheet.getCell(2, 4)).setValue(now);
				style = book.createCellStyle(true);
				style.setDataFormat("h:mm AM/PM;@");
				cell.setCellStyle(style);
				
				//
				(cell = sheet.getCell(3, 0)).setValue("Formula:");
				(cell = sheet.getCell(3, 1)).setNumberValue(1D);
				(cell = sheet.getCell(3, 2)).setNumberValue(2D);
				(cell = sheet.getCell(3, 3)).setNumberValue(3D);
				(cell = sheet.getCell(3, 4)).setFormulaValue("SUM(B4:D4)");
				
				(cell = sheet.getCell(4, 0)).setStringValue("this is a long long long long long string");
				
				(cell = sheet.getCell(5, 0)).setStringValue("merege A6:C6");
				sheet.addMergedRegion(new CellRegion(5,0,5,2));
				(cell = sheet.getCell(5, 3)).setStringValue("merege D6:E7");
				sheet.addMergedRegion(new CellRegion(5,3,6,4));
				
				
				cell = sheet.getCell(9, 6);
				cell.setStringValue("G9");
				style = book.createCellStyle(true);
				style.setFont(book.createFont(true));
				style.getFont().setColor(book.createColor("#FF0000"));
				style.getFont().setHeightPoints(16);
				style.setFillPattern(FillPattern.SOLID_FOREGROUND);
				style.setFillColor(book.createColor("#AAAAAA"));
				
				
				sheet.getColumn(6).setWidth(150);
				sheet.getRow(9).setHeight(100);
				
				style.setAlignment(Alignment.RIGHT);
				style.setVerticalAlignment(VerticalAlignment.CENTER);
				
				style.setBorderTop(BorderType.THIN);
				style.setBorderBottom(BorderType.THIN);
				style.setBorderLeft(BorderType.THIN);
				style.setBorderRight(BorderType.THIN);
				style.setBorderTopColor(book.createColor("#FF0000"));
				style.setBorderBottomColor(book.createColor("#FFFF00"));
				style.setBorderLeftColor(book.createColor("#FF00FF"));
				style.setBorderRightColor(book.createColor("#00FFFF"));
				
				cell.setCellStyle(style);
				
				//row/column style
				
				style = book.createCellStyle(true);
				style.setFillPattern(FillPattern.SOLID_FOREGROUND);
				style.setFillColor(book.createColor("#FFAAAA"));
				sheet.getRow(17).setCellStyle(style);
				sheet.getCell(17, 0).setStringValue("row style");
				style = book.createCellStyle(true);
				style.setFillPattern(FillPattern.SOLID_FOREGROUND);
				style.setFillColor(book.createColor("#AAFFAA"));
				sheet.getColumn(17).setCellStyle(style);
				sheet.getColumn(17).setWidth(100);
				sheet.getCell(0, 17).setStringValue("column style");
				
				
				
				sheet.addPicture(Format.JPG, getTestImageData(), new NViewAnchor(12, 3, 30, 5, 600, 300));
				
				
				sheet = book.createSheet("Sheet 2");
				sheet.getCell(0, 0).setValue("=SUM('Sheet 1'!B4:D4)");
			}


			private byte[] getTestImageData() {
				InputStream is = null;
				try{
					is = getClass().getResourceAsStream("test.jpg");
					ByteArrayOutputStream os = new ByteArrayOutputStream();
					byte[] b = new byte[1024];
					int r;
					while( (r = is.read(b)) !=-1){
						os.write(b,0,r);
					}
					return os.toByteArray();
				} catch (IOException e) {
					throw new RuntimeException(e.getMessage(),e);
				}finally{
					if(is!=null){
						try {
							is.close();
						} catch (IOException e) {}
					}
				}
			}
		};
	}

}
