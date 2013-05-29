package zss.testapp.issue;


import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zkmax.zul.Filedownload;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.model.impl.pdf.PdfExporter;
import org.zkoss.zss.ui.Spreadsheet;

public class ExportComposer195 extends SelectorComposer<Component> {

	private static final long serialVersionUID = 1L;
	
	Exporter pdfExporter = Exporters.getExporter("pdf");
	
	@Wire("spreadsheet")
	private Spreadsheet spreadsheet;


	@Listen("onClick = button[label='Export PDF']")
	public void toPdf() throws IOException{

		File file = new File("exported.pdf");
		file.deleteOnExit();
		FileOutputStream fos = new FileOutputStream(file);
		spreadsheet.getSelectedXSheet().setPrintGridlines(true);
		pdfExporter.export(spreadsheet.getSelectedSheet(), fos);
		Filedownload.save(file, "application/pdf");
	}
	
	@Listen("onClick = button[label='Export selection to PDF']")
	public void selection2Pdf() throws IOException{

		File file = new File("exported.pdf");
		file.deleteOnExit();
		FileOutputStream fos = new FileOutputStream(file);
//		((PdfExporter)pdfExporter).enableGridLines(true);
		spreadsheet.getSelectedXSheet().setPrintGridlines(true);
		pdfExporter.export(spreadsheet.getSelectedSheet(), spreadsheet.getSelection(), fos);
		Filedownload.save(file, "application/pdf");
	}
	
}
