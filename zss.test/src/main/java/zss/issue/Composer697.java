/**
 * This is about DemoComposer
 * @author mqiu
 * @date 10/27/2014
 * 
 * Description: a demo to show the formula issue
 * If a cell's formula references another formula cell and the second cell is
 * not preloaded, the formula value would be incorrect. The second cell would not 
 * calculate even after loaded.
 */

package zss.issue;

import java.awt.Desktop;
import java.awt.Desktop.Action;
import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;

import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.impl.pdf.PdfExporter;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Button;
import org.zkoss.zul.Window;

public class Composer697 extends SelectorComposer<Window> {

	private static final long serialVersionUID = -7360829196117880724L;
	@Wire
	Spreadsheet ss;
	
	@Wire
	Button export;

	@Listen("onClick=#export")
	public void exportHeaderFooter() {
		Book book = ss.getBook();
		
		File temp = getTempFile("Issue697ChartsEngine",".pdf");
		
		exportBook(book.getInternalBook(), temp);
		
		open(temp);
	}
	
	private void exportBook(SBook book, File file) {
		
		PdfExporter exporter = new PdfExporter();
		try {
			exporter.export(book, file);
		} catch (IOException e) {
			e.printStackTrace();
		}

	}
	
	private synchronized File getTempFile(String prefix,String postfix){
		File tempFolder = new File(System.getProperty("java.io.tmpdir"));
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmssSSS");
		File file = null;
		do{
			file = new File(tempFolder,prefix+"-"+sdf.format(new java.util.Date())+postfix);
			try {
				Thread.sleep(10);
			} catch (InterruptedException e) {}
		}while(file.exists());
		return file;
	}

	private void open(File file) {
		if (Desktop.isDesktopSupported()) {
			if (Desktop.getDesktop().isSupported(Action.OPEN)) {
				try {
					Desktop.getDesktop().open(file);
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
}
