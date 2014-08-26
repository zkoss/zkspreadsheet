package zss.testapp;


import static org.zkoss.zss.api.CellOperationUtil.*;
import static org.zkoss.zss.api.Ranges.range;

import java.io.*;

import org.zkoss.image.AImage;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.*;
import org.zkoss.zss.api.*;
import org.zkoss.zss.api.Range.ApplyBorderType;
import org.zkoss.zss.api.model.*;
import org.zkoss.zss.api.model.CellStyle.BorderType;
import org.zkoss.zss.api.model.Chart.Grouping;
import org.zkoss.zss.api.model.Chart.LegendPosition;
import org.zkoss.zss.api.model.Chart.Type;
import org.zkoss.zss.api.model.Picture.Format;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Filedownload;

public class ExportComposer extends SelectorComposer<Component> {

	private static final long serialVersionUID = 1L;
	
	Exporter pdfExporter = Exporters.getExporter("pdf");
	Exporter htmlExporter = Exporters.getExporter("html");
	
	@Wire("#source")
	private Spreadsheet srcSpreadsheet;
	@Wire("#destination")
	private Spreadsheet dstSpreadsheet;
	
	/*
	 * input texts with various styles
	 * FIXME if we export to dstSpreadsheet first, then input styled text, it will throw an exception:
	 *  >>	at org.openxmlformats.schemas.spreadsheetml.x2006.main.impl.CTWorksheetImpl.getColsArray(Unknown Source)
>>	at org.zkoss.poi.xssf.usermodel.helpers.ColumnHelper.getOrCreateColumn1Based(ColumnHelper.java:259)
>>	at org.zkoss.poi.xssf.usermodel.helpers.ColumnHelper.setColWidth(ColumnHelper.java:243)
	 */
	@Listen("onClick = #inputText")
	public void inputStyledText(){
		Sheet sheet = srcSpreadsheet.getSelectedSheet();
		int c = 0 ;
		range(sheet, 0, c).setCellEditText("Bold");
		applyFontBoldweight(range(sheet, 0, c),Font.Boldweight.BOLD);
		
		c++;
		range(sheet, 0, c).setCellEditText("Italic");
		applyFontItalic(range(sheet, 0, c), true);
		
		c++;
		range(sheet, 0, c).setCellEditText("Underline");
		applyFontUnderline(range(sheet, 0, c), Font.Underline.SINGLE);
		
		c++;
		range(sheet, 0, c).setCellEditText("Strikeout");
		applyFontStrikeout(range(sheet, 0, c),  true);

		c++;
		range(sheet, 0, c).setCellEditText("font #ff0000");
		applyFontColor(range(sheet, 0, c), "#ff0000");
		range(sheet, 0, c).setColumnWidth(100);
		
		c++;
		range(sheet, 0, c).setCellEditText("background #00ff00");
		applyFillColor(range(sheet, 0, c), "#00ff00");
		range(sheet, 0, c).setColumnWidth(150);
		
		c++;
		range(sheet, 0, c).setCellEditText("center align");
		applyAlignment(range(sheet, 0, c), CellStyle.Alignment.CENTER);
		applyVerticalAlignment(range(sheet, 0, c), CellStyle.VerticalAlignment.TOP);
		range(sheet, 0, c).setColumnWidth(100);
		range(sheet, 0, c).setRowHeight(80);
		
		//input a table
		applyBorder(range(sheet, 2, 0, 6 ,3), ApplyBorderType.FULL, BorderType.THIN, "#000000");
		
		range(sheet, 2, 0).setCellEditText("Browser Market");
		range(sheet, 2,0 ,2,3).merge(false);
		applyFontHeightPoints(range(sheet, 2,0 ,2,3), (short)16);
		
		range(sheet, 3, 0).setCellEditText("Month");
		range(sheet, 3, 1).setCellEditText("IE");
		range(sheet, 3, 2).setCellEditText("Chrome");
		range(sheet, 3, 3).setCellEditText("Firefox");
		
		range(sheet, 4, 0).setCellEditText("Jan");
		range(sheet, 4, 1).setCellEditText("34");
		range(sheet, 4, 2).setCellEditText("26");
		range(sheet, 4, 3).setCellEditText("22");
		
		range(sheet, 5, 0).setCellEditText("Feb");
		range(sheet, 5, 1).setCellEditText("32");
		range(sheet, 5, 2).setCellEditText("27");
		range(sheet, 5, 3).setCellEditText("22");
		
		range(sheet, 6, 0).setCellEditText("Mar");
		range(sheet, 6, 1).setCellEditText("31");
		range(sheet, 6, 2).setCellEditText("28");
		range(sheet, 6, 3).setCellEditText("22");
	}
	

	/**
	 * insert a chart and a picture
	 */
	@Listen("onClick = #inputPicture")
	public void inputChartPicture() throws IOException{
		Sheet sheet = srcSpreadsheet.getSelectedSheet();
		
		
		SheetOperationUtil.addPicture(range(sheet,3, 8 ), 
				new AImage(ExportComposer.class.getResource("zklogo.png")).getByteData(),
				Format.PNG, 100, 100);
		SheetOperationUtil.addPicture(range(sheet,8,  8), 
				new AImage(ExportComposer.class.getResource("zkessentials.png")).getByteData(),
				Format.PNG, 300, 100);

		SheetOperationUtil.addPicture(range(sheet,12, 8 ), 
				new AImage(ExportComposer.class.getResource("zkstudio.png")).getByteData(),
				Format.PNG, 300, 100);

		
	}
	
	@Listen("onClick = #insertChart")
	public void insertChart() {
		Sheet sheet = srcSpreadsheet.getSelectedSheet();
		range(sheet, 0, 0).setCellEditText("series 1");
		range(sheet, 0, 1).setCellEditText("series 2");
		for(int i = 0; i < 10; ++i) {
			range(sheet, i + 1, 0).setCellEditText( i + 3 + "");
			range(sheet, i + 1, 1).setCellEditText( i + 10 + "");
		}
		
		Type[][] types = new Type[][]{{Type.AREA, Type.AREA_3D, Type.BAR, Type.BAR_3D},
				{Type.LINE, Type.LINE_3D, Type.SCATTER}};
		for(int r = 0; r < types.length; ++r) {
			Type[] tt = types[r];
			for(int c = 0; c < tt.length; ++c) {
				Type t = tt[c];
				SheetOperationUtil.addChart(range(sheet, 0, 0, 10, 1),  t, Grouping.STANDARD, LegendPosition.TOP);
			}
		}
		
		types = new Type[][]{{Type.DOUGHNUT, Type.PIE, Type.PIE_3D}};
		for(int r = 0; r < types.length; ++r) {
			Type[] tt = types[r];
			for(int c = 0; c < tt.length; ++c) {
				Type t = tt[c];
//				ChartData cd = ChartDataUtil.getChartData(sheet, new AreaRef(0, 0, 10, 0), t);
				SheetOperationUtil.addChart(range(sheet, 0, 0, 10, 0), t, Grouping.STANDARD, LegendPosition.TOP);
			}
		}
	}
	
	@Listen("onClick = #applyAutoFilter")
	public void applyAutoFilter(){
		Sheet sheet = srcSpreadsheet.getSelectedSheet();
		//FIXME API doesn't work
		SheetOperationUtil.applyAutoFilter(range(sheet,2, 0));
	}
	
	@Listen("onClick = #protect")
	public void protectSheet(){
		Sheet sheet = srcSpreadsheet.getSelectedSheet();
		
		//FIXME API doesn't work
		SheetOperationUtil.protectSheet(range(sheet), null, null);
		
	}
	
	@Listen("onClick = #exportImport")
	public void export2Destination() throws IOException{
		File exportedFile = new File("exported.xlsx");
		FileOutputStream fos = new FileOutputStream(exportedFile);
		//shall get exporter for each exporting
		Exporters.getExporter("excel").export(srcSpreadsheet.getBook(), fos);
		fos.flush();
		fos.close();


		//import
		Importer importer = Importers.getImporter("excel");
		Book book = importer
				.imports(new FileInputStream(exportedFile), "exported");
		dstSpreadsheet.setBook(book);

		exportedFile.delete();
	}

	@Listen("onClick = button[label='Export Excel']")
	public void toExcel() throws IOException{
		File file = new File("exported.xlsx");
		FileOutputStream fos = new FileOutputStream(file);
		//shall get exporter for each exporting
		Exporters.getExporter("excel").export(srcSpreadsheet.getBook(), fos);
		Filedownload.save(file, "application/excel");
	}

	@Listen("onClick = button[label='Export PDF']")
	public void toPdf() throws IOException{

		File file = new File("exported.pdf");
		FileOutputStream fos = new FileOutputStream(file);
		pdfExporter.export(srcSpreadsheet.getBook(), fos);
		Filedownload.save(file, "application/pdf");
	}

	@Listen("onClick = button[label='Export HTML']")
	public void toHtml() throws IOException{
		File file = new File("exported.html");
		FileOutputStream fos = new FileOutputStream(file);
		htmlExporter.export(srcSpreadsheet.getBook(), fos);
		Filedownload.save(file, "text/html");
	}
}
