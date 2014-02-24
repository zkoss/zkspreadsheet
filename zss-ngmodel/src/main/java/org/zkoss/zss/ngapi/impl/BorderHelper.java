package org.zkoss.zss.ngapi.impl;

import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngapi.NRange.ApplyBorderType;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NColor;
import org.zkoss.zss.ngmodel.NSheet;

public class BorderHelper extends RangeHelperBase {

	static final short TOP = 0x01;
	static final short BOTTOM = 0x02;
	static final short LEFT = 0x04;
	static final short RIGHT = 0x08;
	
	final int maxRowIndex;
	final int maxColumnIndex;
	final static String BLACK = "#000000";
	
	public BorderHelper(NRange range) {
		super(range);
		NBook book = range.getSheet().getBook();
		maxRowIndex = book.getMaxRowIndex();
		maxColumnIndex = book.getMaxColumnIndex();
	}

	public void applyBorder(ApplyBorderType borderType, BorderType lineStyle,
			String borderColor) {
		if(borderType == ApplyBorderType.DIAGONAL || borderType==ApplyBorderType.DIAGONAL_DOWN ||
				borderType == ApplyBorderType.DIAGONAL_UP){
			//not implement, just return;
		}
		NSheet sheet = range.getSheet();
		int row = getRow();
		int column = getColumn();
		int lastRow = getLastRow();
		int lastColumn = getLastColumn();
		for(int r = row ; r<= lastRow;r++){
			for (int c = column ; c<= lastColumn; c++){
				short location = 0; 
				if(r==row){
					location |= TOP;
				}
				if(r==lastRow){
					location |= BOTTOM;
				}
				if(c==column){
					location |= LEFT;
				}
				if(c==lastColumn){
					location |= RIGHT;
				}
				handleCellBorder(sheet,r,c,borderType,lineStyle,borderColor,location);
			}
		}
		
	}

	private void handleCellBorder(NSheet sheet, int r, int c,
			ApplyBorderType borderType, BorderType lineStyle,
			String borderColor, short location) {
		switch(borderType){
		case FULL:
			handleFull(sheet,r,c,lineStyle,borderColor,location);
			break;
		case EDGE_BOTTOM:
			handleEdgeBottom(sheet,r,c,lineStyle,borderColor,location);
			break;
		case EDGE_RIGHT:
			handleEdgeRight(sheet,r,c,lineStyle,borderColor,location);
			break;
		case EDGE_TOP:
			handleEdgeTop(sheet,r,c,lineStyle,borderColor,location);
			break;
		case EDGE_LEFT:
			handleEdgeLeft(sheet,r,c,lineStyle,borderColor,location);
			break;
		case OUTLINE:
			handleOutline(sheet,r,c,lineStyle,borderColor,location);
			break;
		case INSIDE:
			handleInside(sheet,r,c,lineStyle,borderColor,location);
			break;
		case INSIDE_HORIZONTAL:
			handleInsideHorizontal(sheet,r,c,lineStyle,borderColor,location);
			break;
		case INSIDE_VERTICAL:
			handleInsideVertical(sheet,r,c,lineStyle,borderColor,location);
			break;
		}
	}

	private void handleInside(NSheet sheet, int r, int c, BorderType lineStyle,
			String borderColor, short location) {
		//ignore when RIGHT or BOTTOM
		//right and bottom border
		short at = StyleUtil.BORDER_EDGE_LEFT | StyleUtil.BORDER_EDGE_TOP | StyleUtil.BORDER_EDGE_RIGHT | StyleUtil.BORDER_EDGE_BOTTOM;
		//top border when top
		if((location&TOP)!=0){
			resetBorderBottom(sheet,r-1,c,lineStyle,borderColor);
			at -= StyleUtil.BORDER_EDGE_TOP;
		}
		//left border when left
		if((location&LEFT)!=0){
			resetBorderRight(sheet,r,c-1,lineStyle,borderColor);
			at -= StyleUtil.BORDER_EDGE_LEFT;
		}
		
		//right and bottom border
		if((location&BOTTOM)!=0){
			resetBorderTop(sheet,r+1,c,lineStyle,borderColor);
			at -= StyleUtil.BORDER_EDGE_BOTTOM;
			
		}
		if((location&RIGHT)!=0){
			resetBorderLeft(sheet,r,c+1,lineStyle,borderColor);
			at -= StyleUtil.BORDER_EDGE_RIGHT;
		}
		
		StyleUtil.setBorder(sheet, r, c, borderColor, lineStyle,at);	
		
	}
	
	private void handleInsideVertical(NSheet sheet, int r, int c, BorderType lineStyle,
			String borderColor, short location) {
		//ignore when RIGHT or BOTTOM
		//right border
		short at = StyleUtil.BORDER_EDGE_LEFT | StyleUtil.BORDER_EDGE_RIGHT;
		
		if((location&LEFT)!=0){
			at -= StyleUtil.BORDER_EDGE_LEFT;
		}
		if((location&RIGHT)!=0){
			at -= StyleUtil.BORDER_EDGE_RIGHT;
		}
		
		StyleUtil.setBorder(sheet, r, c, borderColor, lineStyle,at);
	}
	
	private void handleInsideHorizontal(NSheet sheet, int r, int c, BorderType lineStyle,
			String borderColor, short location) {
		//ignore when RIGHT or BOTTOM
		//bottom border
			
		short at = StyleUtil.BORDER_EDGE_BOTTOM | StyleUtil.BORDER_EDGE_TOP;
		
		if((location&BOTTOM)!=0){
			at -= StyleUtil.BORDER_EDGE_BOTTOM;
		}
		if((location&TOP)!=0){
			at -= StyleUtil.BORDER_EDGE_TOP;
		}
		
		StyleUtil.setBorder(sheet, r, c, borderColor, lineStyle,at);
	}

	private void handleOutline(NSheet sheet, int r, int c,
			BorderType lineStyle, String borderColor, short location) {
		//bottom border when BOTTOM
		//right border when RIGHT
		//top border when TOP
		//right border when RIGHT
		short at = 0;
		if((location&TOP)!=0){
			resetBorderBottom(sheet,r-1,c,lineStyle,borderColor);
			at |= StyleUtil.BORDER_EDGE_TOP;
		}
		if((location&LEFT)!=0){
			resetBorderRight(sheet,r,c-1,lineStyle,borderColor);
			at |= StyleUtil.BORDER_EDGE_LEFT;
		}
		if((location&BOTTOM)!=0){
			resetBorderTop(sheet,r+1,c,lineStyle,borderColor);
			at |= StyleUtil.BORDER_EDGE_BOTTOM;
		}
		if((location&RIGHT)!=0){
			resetBorderLeft(sheet,r,c+1,lineStyle,borderColor);
			at |= StyleUtil.BORDER_EDGE_RIGHT;
		}
		StyleUtil.setBorder(sheet, r, c, borderColor, lineStyle,at);
		
	}

	private void handleEdgeLeft(NSheet sheet, int r, int c,
			BorderType lineStyle, String borderColor, short location) {
		//left border when LEFT
		if((location&LEFT)!=0){
			resetBorderRight(sheet,r,c-1,lineStyle,borderColor);
			StyleUtil.setBorder(sheet, r, c, borderColor, lineStyle,StyleUtil.BORDER_EDGE_LEFT);
		}
	}

	private void handleEdgeTop(NSheet sheet, int r, int c,
			BorderType lineStyle, String borderColor, short location) {
		//top border when TOP
		if((location&TOP)!=0){
			resetBorderBottom(sheet,r-1,c,lineStyle,borderColor);
			StyleUtil.setBorder(sheet, r, c, borderColor, lineStyle,StyleUtil.BORDER_EDGE_TOP);
		}
	}

	private void handleEdgeRight(NSheet sheet, int r, int c,
			BorderType lineStyle, String borderColor, short location) {
		//right border when RIGHT
		if((location&RIGHT)!=0){
			resetBorderLeft(sheet,r,c+1,lineStyle,borderColor);
			StyleUtil.setBorder(sheet, r, c, borderColor, lineStyle,StyleUtil.BORDER_EDGE_RIGHT);
		}
	}

	private void handleEdgeBottom(NSheet sheet, int r, int c,
			BorderType lineStyle, String borderColor, short location) {
		//bottom border when BOTTOM
		if((location&BOTTOM)!=0){
			resetBorderTop(sheet,r+1,c,lineStyle,borderColor);
			StyleUtil.setBorder(sheet, r, c, borderColor, lineStyle,StyleUtil.BORDER_EDGE_BOTTOM);
		}
		
	}

	private void handleFull(NSheet sheet, int r, int c,BorderType lineStyle,
			String borderColor, short location) {
		short at = 0;
		//top border when top
		if((location&TOP)!=0){
			resetBorderBottom(sheet,r-1,c,lineStyle,borderColor);
		}
		//left border when left
		if((location&LEFT)!=0){
			resetBorderRight(sheet,r,c-1,lineStyle,borderColor);
		}
		
		//right and bottom border
		if((location&BOTTOM)!=0){
			resetBorderTop(sheet,r+1,c,lineStyle,borderColor);
			
		}
		if((location&RIGHT)!=0){
			resetBorderLeft(sheet,r,c+1,lineStyle,borderColor);
		}
		at |= StyleUtil.BORDER_EDGE_LEFT;
		at |= StyleUtil.BORDER_EDGE_TOP;
		at |= StyleUtil.BORDER_EDGE_RIGHT;
		at |= StyleUtil.BORDER_EDGE_BOTTOM;
		StyleUtil.setBorder(sheet, r, c, borderColor, lineStyle,at);
	}

	private boolean isOutOfBoundCell(NSheet sheet, int r, int c){
		return (r<0|| c<0 || r>maxRowIndex || c>maxColumnIndex);
	}
	
	private boolean resetBorderTop(NSheet sheet,int r, int c,BorderType borderType,String borderColor){
		if(isOutOfBoundCell(sheet,r,c)){
			return false;
		}
		
		NCell cell = sheet.getCell(r,c);
		if(cell.isNull())
			return false;
		
		NCellStyle style = cell.getCellStyle();
		BorderType bType = style.getBorderTop();
		String bColor = style.getBorderTopColor().getHtmlColor();
		if(NCellStyle.BorderType.NONE.equals(bType)){
			return false;
		}
		//same border
		if(borderType.equals(bType) && borderColor.equals(bColor)){
			return false;
		}
		StyleUtil.setBorder(sheet, r, c,BLACK, NCellStyle.BorderType.NONE,StyleUtil.BORDER_EDGE_TOP);
		return true;
	}
	private boolean resetBorderBottom(NSheet sheet,int r, int c,BorderType borderType,String borderColor){
		if(isOutOfBoundCell(sheet,r,c)){
			return false;
		}
		
		NCell cell = sheet.getCell(r,c);
		if(cell.isNull())
			return false;
		
		NCellStyle style = cell.getCellStyle();
		BorderType bType = style.getBorderBottom();
		String bColor = style.getBorderBottomColor().getHtmlColor();
		if(NCellStyle.BorderType.NONE.equals(bType)){
			return false;
		}
		//same border
		if(borderType.equals(bType) && borderColor.equals(bColor)){
			return false;
		}
		
		StyleUtil.setBorder(sheet, r, c,BLACK, NCellStyle.BorderType.NONE,StyleUtil.BORDER_EDGE_BOTTOM);
		
		return true;
	}
	private boolean resetBorderLeft(NSheet sheet,int r, int c,BorderType borderType,String borderColor){
		if(isOutOfBoundCell(sheet,r,c)){
			return false;
		}
		
		NCell cell = sheet.getCell(r,c);
		if(cell.isNull())
			return false;
		
		NCellStyle style = cell.getCellStyle();
		BorderType bType = style.getBorderLeft();
		String bColor = style.getBorderLeftColor().getHtmlColor();
		if(NCellStyle.BorderType.NONE.equals(bType)){
			return false;
		}
		//same border
		if(borderType.equals(bType) && borderColor.equals(bColor)){
			return false;
		}
		StyleUtil.setBorder(sheet, r, c,BLACK, NCellStyle.BorderType.NONE,StyleUtil.BORDER_EDGE_LEFT);
		return true;
	}
	private boolean resetBorderRight(NSheet sheet,int r, int c,BorderType borderType,String borderColor){
		if(isOutOfBoundCell(sheet,r,c)){
			return false;
		}
		
		NCell cell = sheet.getCell(r,c);
		if(cell.isNull())
			return false;
		
		NCellStyle style = cell.getCellStyle();
		BorderType bType = style.getBorderRight();
		String bColor = style.getBorderRightColor().getHtmlColor();
		if(NCellStyle.BorderType.NONE.equals(bType)){
			return false;
		}
		//same border
		if(borderType.equals(bType) && borderColor.equals(bColor)){
			return false;
		}
		StyleUtil.setBorder(sheet, r, c,BLACK, NCellStyle.BorderType.NONE,StyleUtil.BORDER_EDGE_RIGHT);
		return true;
	}
}
