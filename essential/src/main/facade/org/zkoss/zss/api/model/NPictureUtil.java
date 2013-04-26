package org.zkoss.zss.api.model;

import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.zss.api.NSheetAnchor;
import org.zkoss.zss.api.UnitUtil;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Worksheet;

public class NPictureUtil {

	
	public static NSheetAnchor toFilledAnchor(NSheet sheet,int row, int column, int widthPx, int heightPx){
		int lRow = 0;
		int lColumn = 0;
		int lX = 0;
		int lY = 0;
		
		Worksheet ws = sheet.getNative();
		Book book = ws.getBook();
		for(int i = column;;i++){
			if(ws.isColumnHidden(i)){
				continue;
			}
			int wPx = UnitUtil.getColumnWidthInPx(sheet,i);
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
			
			int hPx = UnitUtil.getRowHeightInPx(sheet, i);
			heightPx -= hPx;
			if(heightPx<=0){
				lRow = i-1;
				lY = hPx + heightPx;
				break;
			}
		}
		return new NSheetAnchor(row,column,0,0,lRow,lColumn,lX,lY);
	}
}
