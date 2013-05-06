package org.zkoss.zss.api;

import org.zkoss.image.AImage;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Chart;
import org.zkoss.zss.api.model.ChartData;
import org.zkoss.zss.api.model.Picture.Format;

public class SheetOperationUtil {

	public static void toggleAutoFilter(Range range) {
		if(range.isProtected())
			return;
		range.enableAutoFilter(!range.isAutoFilterEnabled());
	}
	
	public static void resetAutoFilter(Range range) {
		if(range.isProtected())
			return;
		range.resetAutoFilter();
	}
	
	public static void applyAutoFilter(Range range) {
		if(range.isProtected())
			return;
		range.applyAutoFilter();
	}
	
	public static void addPicture(Range range, AImage image){
		addPicture(range,image.getByteData(),getPictureFormat(image),image.getWidth(),image.getHeight());
	}
	
	public static void addPicture(Range range, byte[] binary, Format format,int widthPx, int heightPx){
		SheetAnchor anchor = UnitUtil.toFilledAnchor(range.getSheet(), range.getRow(),range.getColumn(),
				widthPx, heightPx);
		addPicture(range,anchor,binary,format);
		
	}
	public static void addPicture(Range range, SheetAnchor anchor, byte[] binary, Format format){
		if(range.isProtected())
			return;
		range.addPicture(anchor, binary, format);
	}
	
	public static Format getPictureFormat(AImage image) {
		String format = image.getFormat();
		if ("dib".equalsIgnoreCase(format)) {
			return Format.DIB;
		} else if ("emf".equalsIgnoreCase(format)) {
			return Format.EMF;
		} else if ("wmf".equalsIgnoreCase(format)) {
			return Format.WMF;
		} else if ("jpeg".equalsIgnoreCase(format)) {
			return Format.JPEG;
		} else if ("pict".equalsIgnoreCase(format)) {
			return Format.PICT;
		} else if ("png".equalsIgnoreCase(format)) {
			return Format.PNG;
		}
		return null;
	}
	
	
	public static void addChart(Range range, ChartData data, Chart.Type type, Chart.Grouping grouping,
			Chart.LegendPosition pos) {
		SheetAnchor anchor = toChartAnchor(range);
		addChart(range,anchor, data, type, grouping, pos);
	}
	
	public static void addChart(Range range, SheetAnchor anchor, ChartData data, Chart.Type type, Chart.Grouping grouping,
			Chart.LegendPosition pos) {
		if (range.isProtected())
			return;
		range.addChart(anchor, data, type, grouping, pos);
	}
	
	
	public static SheetAnchor toChartAnchor(Range range) {
		int row = range.getRow();
		int col = range.getColumn();
		int lRow = range.getLastRow();
		int lCol = range.getLastColumn();
		int w = lCol-col+1;
		//shift 2 column right for the selection width 
		return new SheetAnchor(row, lCol+2, 
				row==lRow?row+7:lRow+1, col==lCol?lCol+7+w:lCol+2+w);
	}

	public static void protectSheet(Range range, String password,String newpasswrod) {
		//TODO the spec?
//		if (range.isProtected())
//			return;
		
		range.protectSheet(newpasswrod);
	}

	public static void displaySheetGridlines(Range range,boolean enable) {
		range.setDisplaySheetGridlines(enable);
	}

	public static void addSheet(Range range, final String prefix) {
		range.sync(new RangeRunner(){
			public void run(Range range) {
				String name;
				Book book = range.getBook();
				int numSheet = book.getNumberOfSheets();
				do{
					numSheet++;
					name = prefix+numSheet;
				}while(book.getSheet(name)!=null);
				range.createSheet(name);
			}});
	}
	public static void createSheet(Range range, final String name) {
		range.sync(new RangeRunner(){
			public void run(Range range) {
				Book book = range.getBook();
				if(book.getSheet(name)!=null){
					throw new IllegalArgumentException("another sheet has same name already : "+name);
				}
				//it is possible throw a exception if there is another sheet has same name
				range.createSheet(name);
			}});
	}

	public static void renameSheet(Range range, final String newname) {
		range.sync(new RangeRunner() {
			public void run(Range range) {
				Book book = range.getBook();

				if (book.getSheet(newname) != null) {
					return;
				}
				range.setSheetName(newname);
			}
		});
	}

	public static void setSheetOrder(Range range, int pos) {
		range.setSheetOrder(pos);
	}

	public static void deleteSheet(Range range) {
		range.sync(new RangeRunner() {
			public void run(Range range) {
				int num = range.getBook().getNumberOfSheets();
				if (num <= 1) {
					// don't do it
					return;
				}
				range.deleteSheet();
			}
		});

	}
}
