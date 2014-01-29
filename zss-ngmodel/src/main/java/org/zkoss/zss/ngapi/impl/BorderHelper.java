package org.zkoss.zss.ngapi.impl;

import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngapi.NRange.ApplyBorderType;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NSheet;

public class BorderHelper extends RangeHelperBase {

	static final short TOP = 0x01;
	static final short BOTTOM = 0x02;
	static final short LEFT = 0x04;
	static final short RIGHT = 0x08;
	
	final int maxRowIndex;
	final int maxColumnIndex;
	
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
		short at = StyleUtil.BORDER_EDGE_RIGHT | StyleUtil.BORDER_EDGE_BOTTOM;
		
		if((location&BOTTOM)!=0){
			at -= StyleUtil.BORDER_EDGE_BOTTOM;
		}
		if((location&RIGHT)!=0){
			at -= StyleUtil.BORDER_EDGE_RIGHT;
		}
		
		StyleUtil.setBorder(sheet, r, c, borderColor, lineStyle,at);	
		
	}
	
	private void handleInsideVertical(NSheet sheet, int r, int c, BorderType lineStyle,
			String borderColor, short location) {
		//ignore when RIGHT or BOTTOM
		//right border
		short at = StyleUtil.BORDER_EDGE_RIGHT;

		if((location&RIGHT)!=0){
			return;
		}
		
		StyleUtil.setBorder(sheet, r, c, borderColor, lineStyle,at);
	}
	
	private void handleInsideHorizontal(NSheet sheet, int r, int c, BorderType lineStyle,
			String borderColor, short location) {
		//ignore when RIGHT or BOTTOM
		//bottom border
			
		short at = StyleUtil.BORDER_EDGE_BOTTOM;
		
		if((location&BOTTOM)!=0){
			return;
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
			if(r-1>=0){
				StyleUtil.setBorder(sheet, r-1, c, borderColor, lineStyle,StyleUtil.BORDER_EDGE_BOTTOM);
			}
			at |= StyleUtil.BORDER_EDGE_TOP;
		}
		if((location&LEFT)!=0){
			if(c-1>=0){
				StyleUtil.setBorder(sheet, r, c-1, borderColor, lineStyle,StyleUtil.BORDER_EDGE_RIGHT);
			}
			at |= StyleUtil.BORDER_EDGE_LEFT;
		}
		if((location&BOTTOM)!=0){
			if(r+1<=maxRowIndex){
				StyleUtil.setBorder(sheet, r+1, c, borderColor, lineStyle,StyleUtil.BORDER_EDGE_TOP);
			}
			at |= StyleUtil.BORDER_EDGE_BOTTOM;
		}
		if((location&RIGHT)!=0){
			if(c+1<=maxColumnIndex){
				StyleUtil.setBorder(sheet, r, c+1, borderColor, lineStyle,StyleUtil.BORDER_EDGE_LEFT);
			}
			at |= StyleUtil.BORDER_EDGE_RIGHT;
		}
		StyleUtil.setBorder(sheet, r, c, borderColor, lineStyle,at);
		
	}

	private void handleEdgeLeft(NSheet sheet, int r, int c,
			BorderType lineStyle, String borderColor, short location) {
		//left border when LEFT
		if((location&LEFT)!=0){
			if(c-1>=0){
				StyleUtil.setBorder(sheet, r, c-1, borderColor, lineStyle,StyleUtil.BORDER_EDGE_RIGHT);
			}
			StyleUtil.setBorder(sheet, r, c, borderColor, lineStyle,StyleUtil.BORDER_EDGE_LEFT);
		}
	}

	private void handleEdgeTop(NSheet sheet, int r, int c,
			BorderType lineStyle, String borderColor, short location) {
		//top border when TOP
		if((location&TOP)!=0){
			if(r-1>=0){
				StyleUtil.setBorder(sheet, r-1, c, borderColor, lineStyle,StyleUtil.BORDER_EDGE_BOTTOM);
			}
			StyleUtil.setBorder(sheet, r, c, borderColor, lineStyle,StyleUtil.BORDER_EDGE_TOP);
		}
	}

	private void handleEdgeRight(NSheet sheet, int r, int c,
			BorderType lineStyle, String borderColor, short location) {
		//right border when RIGHT
		if((location&RIGHT)!=0){
			if(c+1<=maxColumnIndex){
				StyleUtil.setBorder(sheet, r, c+1, borderColor, lineStyle,StyleUtil.BORDER_EDGE_LEFT);
			}
			StyleUtil.setBorder(sheet, r, c, borderColor, lineStyle,StyleUtil.BORDER_EDGE_RIGHT);
		}
	}

	private void handleEdgeBottom(NSheet sheet, int r, int c,
			BorderType lineStyle, String borderColor, short location) {
		//bottom border when BOTTOM
		if((location&BOTTOM)!=0){
			if(r+1<=maxRowIndex){
				StyleUtil.setBorder(sheet, r+1, c, borderColor, lineStyle,StyleUtil.BORDER_EDGE_TOP);
			}
			StyleUtil.setBorder(sheet, r, c, borderColor, lineStyle,StyleUtil.BORDER_EDGE_BOTTOM);
		}
		
	}

	private void handleFull(NSheet sheet, int r, int c,BorderType lineStyle,
			String borderColor, short location) {
		short at = 0;
		//top border when top
		if((location&TOP)!=0){
			if(r-1>=0){
				StyleUtil.setBorder(sheet, r-1, c, borderColor, lineStyle,StyleUtil.BORDER_EDGE_BOTTOM);
			}
			at |= StyleUtil.BORDER_EDGE_TOP;
		}
		//left border when left
		if((location&LEFT)!=0){
			if(c-1>=0){
				StyleUtil.setBorder(sheet, r, c-1, borderColor, lineStyle,StyleUtil.BORDER_EDGE_RIGHT);
			}
			at |= StyleUtil.BORDER_EDGE_LEFT;
		}
		
		//right and bottom border
		if((location&BOTTOM)!=0){
			if(r+1<=maxRowIndex){
				StyleUtil.setBorder(sheet, r+1, c, borderColor, lineStyle,StyleUtil.BORDER_EDGE_TOP);
			}
			
		}
		if((location&RIGHT)!=0){
			if(c+1<=maxColumnIndex){
				StyleUtil.setBorder(sheet, r, c+1, borderColor, lineStyle,StyleUtil.BORDER_EDGE_LEFT);
			}
		}
		at |= StyleUtil.BORDER_EDGE_RIGHT;
		at |= StyleUtil.BORDER_EDGE_BOTTOM;
		StyleUtil.setBorder(sheet, r, c, borderColor, lineStyle,at);
	}


}
