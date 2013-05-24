package zss.testapp;


import static org.zkoss.zss.api.CellOperationUtil.applyAlignment;
import static org.zkoss.zss.api.CellOperationUtil.applyBackgroundColor;
import static org.zkoss.zss.api.CellOperationUtil.applyBorder;
import static org.zkoss.zss.api.CellOperationUtil.applyFontBoldweight;
import static org.zkoss.zss.api.CellOperationUtil.applyFontColor;
import static org.zkoss.zss.api.CellOperationUtil.applyFontItalic;
import static org.zkoss.zss.api.CellOperationUtil.applyFontSize;
import static org.zkoss.zss.api.CellOperationUtil.applyFontStrikeout;
import static org.zkoss.zss.api.CellOperationUtil.applyFontUnderline;
import static org.zkoss.zss.api.CellOperationUtil.applyVerticalAlignment;
import static org.zkoss.zss.api.Ranges.range;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.zkoss.image.AImage;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zkmax.zul.Filedownload;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range.ApplyBorderType;
import org.zkoss.zss.api.SheetOperationUtil;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.CellStyle.BorderType;
import org.zkoss.zss.api.model.Font;
import org.zkoss.zss.api.model.Picture.Format;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;

public class ExportComposer extends SelectorComposer<Component> {

	private static final long serialVersionUID = 1L;
	
	Exporter excelExporter = Exporters.getExporter("excel");
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
		applyBackgroundColor(range(sheet, 0, c), "#00ff00");
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
		applyFontSize(range(sheet, 2,0 ,2,3), (short)16);
		
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
		excelExporter.export(srcSpreadsheet.getBook(), fos);
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
		excelExporter.export(srcSpreadsheet.getBook(), fos);
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
