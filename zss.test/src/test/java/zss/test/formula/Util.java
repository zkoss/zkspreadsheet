package zss.test.formula;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.zkoss.zss.model.Exporter;
import org.zkoss.zss.model.Exporters;
import org.zkoss.zss.model.Importers;
import org.zkoss.zss.model.Book;

/**
 * a helper for testing
 * 
 * @author kuro, Hawk
 * 
 */
public class Util {

	public static Book loadBook(Object base,String respath) {
		if(base==null){
			base = Util.class;
		}
		if(!(base instanceof Class)){
			base = base.getClass();
		}

		final InputStream is = ((Class)base).getResourceAsStream(respath);
		try {
			return Importers.getImporter("excel").imports(is, respath);
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {
				}
			}
		}
	}

	public static void export(Book workbook, File file) {
		Exporter excelExporter = Exporters.getExporter("Excel");
		OutputStream os = null;
		try {
			os = new FileOutputStream(file);
			excelExporter.export(workbook, os);
		} catch (IOException e) {
			throw new RuntimeException(e.getMessage(),e);
		} finally {
			if(os!=null)
				try {
					os.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
		}
	}
}
