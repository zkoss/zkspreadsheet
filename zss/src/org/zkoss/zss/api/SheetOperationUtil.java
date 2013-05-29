/* SheetOperationUtil.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/5/1 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.api;

import org.zkoss.image.AImage;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.zss.api.Range.AutoFillType;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Chart;
import org.zkoss.zss.api.model.ChartData;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.Picture.Format;
import org.zkoss.zss.api.model.impl.SheetImpl;
import org.zkoss.zss.model.sys.XSheet;

/**
 * The utility to help UI to deal with user's sheet operation of a Range.
 * This utility is the default implementation for handling a spreadsheet, it is also the example for calling {@link Range} APIs 
 * @author dennis
 * @since 3.0.0
 */
public class SheetOperationUtil {

	/**
	 * Toggles autofilter on or off
	 * @param range the range to toggle
	 */
	public static void toggleAutoFilter(Range range) {
		if(range.isProtected())
			return;
		range.enableAutoFilter(!range.isAutoFilterEnabled());
	}
	
	/**
	 * Resets autofilter
	 * @param range the range to reset
	 */
	public static void resetAutoFilter(Range range) {
		if(range.isProtected())
			return;
		range.resetAutoFilter();
	}
	
	/**
	 * Re-apply autofilter
	 * @param range the range to apply
	 */
	public static void applyAutoFilter(Range range) {
		if(range.isProtected())
			return;
		range.applyAutoFilter();
	}
	
	/**
	 * Add picture to the range
	 * @param range the range to add picture to
	 * @param image the image
	 */
	public static void addPicture(Range range, AImage image){
		addPicture(range,image.getByteData(),getPictureFormat(image),image.getWidth(),image.getHeight());
	}
	
	/**
	 * Add picture to the range
	 * @param range the range to add picture
	 * @param binary the image binary data
	 * @param format the image format
	 * @param widthPx the width of image to place
	 * @param heightPx the height of image to place
	 */
	public static void addPicture(Range range, byte[] binary, Format format,int widthPx, int heightPx){
		SheetAnchor anchor = toFilledAnchor(range.getSheet(), range.getRow(),range.getColumn(),
				widthPx, heightPx);
		addPicture(range,anchor,binary,format);
		
	}
	
	/**
	 * Add picture to the range
	 * @param range the range to add picture
	 * @param anchor the picture location
	 * @param binary the image binary data
	 * @param format the image format
	 */
	public static void addPicture(Range range, SheetAnchor anchor, byte[] binary, Format format){
		if(range.isProtected())
			return;
		range.addPicture(anchor, binary, format);
	}
	
	/**
	 * Gets the picture format
	 * @param image the image
	 * @return image format, or null if doens't support
	 */
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
	
	/**
	 * Adds chart to range
	 * @param range the range to add chart
	 * @param data the chart data
	 * @param type the chart type
	 * @param grouping the grouping type
	 * @param pos the legend position type
	 */
	public static void addChart(Range range, ChartData data, Chart.Type type, Chart.Grouping grouping,
			Chart.LegendPosition pos) {
		SheetAnchor anchor = toChartAnchor(range);
		addChart(range,anchor, data, type, grouping, pos);
	}
	
	/**
	 * Adds chart to range
	 * @param range the range to add chart
	 * @param anchor the chart location
	 * @param data the chart data
	 * @param type the chart type
	 * @param grouping the grouping type
	 * @param pos the legend position type
	 */
	public static void addChart(Range range, SheetAnchor anchor, ChartData data, Chart.Type type, Chart.Grouping grouping,
			Chart.LegendPosition pos) {
		if (range.isProtected())
			return;
		range.addChart(anchor, data, type, grouping, pos);
	}
	
	/**
	 * Create a anchor by range's row/column data
	 * @param range the range for chart.
	 * @return a new anchor
	 */
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

	/**
	 * Enables/disables sheet gridlines.
	 * @param range the range to be applied
	 * @param enable true for enable
	 */
	public static void displaySheetGridlines(Range range,boolean enable) {
		range.setDisplaySheetGridlines(enable);
	}

	/**
	 * Add a new sheet to this book
	 * @param range the range to be applied
	 * @param prefix the sheet name prefix, it will selection a new name with this prefix and a counter.
	 */
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
	
	/**
	 * Add a new sheet to this book
	 * @param range the range to be applied
	 * @param name the sheet name, it must not same as another sheet name in this book
	 */
	public static void createSheet(Range range, final String name) {
		range.sync(new RangeRunner(){
			public void run(Range range) {
				Book book = range.getBook();
				if(book.getSheet(name)!=null){
					// don't do it
					return;
				}
				//it is possible throw a exception if there is another sheet has same name
				range.createSheet(name);
			}});
	}

	/**
	 * Rename a sheet
	 * @param range the range to be applied
	 * @param newname the new name of the sheet, it must not same as another sheet name in this book
	 */
	public static void renameSheet(Range range, final String newname) {
		range.sync(new RangeRunner() {
			public void run(Range range) {
				if(range.getSheetName().equals(newname)){
					return;
				}
				
				Book book = range.getBook();
				if (book.getSheet(newname) != null) {
					// don't do it
					return;
				}
				range.setSheetName(newname);
			}
		});
	}

	/**
	 * Sets the sheet order
	 * @param range the range to be applied
	 * @param pos the sheet position
	 */
	public static void setSheetOrder(Range range, int pos) {
		range.setSheetOrder(pos);
	}

	/**
	 * Deletes the sheet, Note you can't delete last sheet
	 * @param range the range to be applied
	 */
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

	/**
	 * According to source range, fills data to destination range automatically
	 * @param src the source range
	 * @param dest the destination range
	 * @param type the fill type
	 */
	public static void autoFill(Range src, Range dest, AutoFillType type) {
		if(dest.isProtected())
			return;
		src.autoFill(dest, type);
	}
	
	
	public static SheetAnchor toFilledAnchor(Sheet sheet,int row, int column, int widthPx, int heightPx){
		int lRow = 0;
		int lColumn = 0;
		int lX = 0;
		int lY = 0;
		
		XSheet ws = ((SheetImpl)sheet).getNative();
//		Book book = ws.getBook();
		for(int i = column;;i++){
			if(ws.isColumnHidden(i)){
				continue;
			}
			int wPx = sheet.getColumnWidth(i);
			widthPx -= wPx;
			if(widthPx<=0){
				lColumn = i-1;
				lX = wPx + widthPx;//offset
				break;
			}
		}
		
		
		for(int i = row;;i++){
			Row srow = ws.getRow(i);
			if(srow!=null && srow.getZeroHeight()){
				continue;
			}
			
			int hPx = sheet.getRowHeight(i);
			heightPx -= hPx;
			if(heightPx<=0){
				lRow = i-1;
				lY = hPx + heightPx;
				break;
			}
		}
		return new SheetAnchor(row,column,0,0,lRow,lColumn,lX,lY);
	}
}
