package org.zkoss.zss.ngmodel.impl;

import java.util.List;

import org.zkoss.zss.ngapi.impl.StyleUtil;
import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.InvalidateModelOpException;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NComment;
import org.zkoss.zss.ngmodel.NDataValidation;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.PasteOption;
import org.zkoss.zss.ngmodel.PasteOption.PasteType;
import org.zkoss.zss.ngmodel.SheetRegion;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.sys.formula.FormulaParseContext;
import org.zkoss.zss.ngmodel.util.Validations;

/*package*/ class CopyCellHelper {

	private final NSheet destSheet;
	private final NBook book;
	private final NCellStyle defaultStyle;
	public CopyCellHelper(NSheet destSheet){
		this.destSheet = destSheet;
		this.book = destSheet.getBook();
		defaultStyle = book.getDefaultCellStyle();
	}
	
	static class CellBuffer{
		final int row;
		final int column;
		CellType type;
		Object value;
		String formula;
		NCellStyle style;
		
		NComment comment;
		NDataValidation validation;
		
		public CellBuffer(int row, int column){
			this.row = row;
			this.column = column;
		}
		public CellType getType() {
			return type;
		}
		public void setType(CellType type) {
			this.type = type;
		}
		
		public int getRow() {
			return row;
		}
		public int getColumn() {
			return column;
		}
		public Object getValue() {
			return value;
		}
		public void setValue(Object value) {
			this.value = value;
		}
		public String getFormula() {
			return formula;
		}
		public void setFormula(String formula) {
			this.formula = formula;
		}
		public NCellStyle getStyle() {
			return style;
		}
		public void setStyle(NCellStyle style) {
			this.style = style;
		}
		public NComment getComment() {
			return comment;
		}
		public void setComment(NComment comment) {
			this.comment = comment;
		}
		public NDataValidation getValidation() {
			return validation;
		}
		public void setValidation(NDataValidation validation) {
			this.validation = validation;
		}
	}
	
	public void copyCell(SheetRegion src, CellRegion dest, PasteOption option) {
		Validations.argNotNull(src);
		Validations.argNotNull(dest);
		NSheet srcSheet = src.getSheet();
		boolean sameSheet = srcSheet == destSheet;
		if(!sameSheet && srcSheet.getBook()!= book){
			throw new IllegalArgumentException("the src sheet must be in the same book");
		}
		if(option==null){
			option = new PasteOption();
		}
		
		if(option.getPasteType()==PasteType.COLUMN_WIDTHS){
			copyColumnWidth(src,dest);
			return;
		}
		
		
		int rowOffset = dest.getRow() - src.getRow();
		int columnOffset = dest.getColumn() - src.getColumn();
		
		int srcColCount = src.getColumnCount();
		int srcRowCount = src.getRowCount();
		int destColCount = dest.getColumnCount();
		int destRowCount = dest.getRowCount();
		
		
		boolean wrongMultiple = (destRowCount>1 && destRowCount%srcRowCount!=0)||(destColCount>1 && destColCount%srcColCount!=0);
		
		int rowMultiple = destRowCount<=1||wrongMultiple?1:destRowCount/srcRowCount;
		int colMultiple = destColCount<=1||wrongMultiple?1:destColCount/srcColCount;
		
		CellRegion finalRegion = new CellRegion(dest.getRow(),dest.getColumn(),
				dest.getRow()+srcRowCount*rowMultiple-1,dest.getColumn()+srcColCount*colMultiple-1);
		
		List<CellRegion> merged = destSheet.getOverlapsMergedRegions(finalRegion); 
		if(merged.size()>0){
			throw new InvalidateModelOpException("Can't paste to "+finalRegion.getReferenceString()+", it overlaps merged cells "+merged.get(0).getReferenceString());
		}
		
		CellBuffer[][] srcBuffer = prepareBuffer(src,option);
		for(int i=0;i<rowMultiple;i++){
			for(int j=0;j<colMultiple;j++){
				int rowMultpleOffset = i*srcRowCount;
				int colMultipleOffset = j*srcColCount;
				CellRegion destRegion = new CellRegion(dest.getRow()+rowMultpleOffset,dest.getColumn()+colMultipleOffset,
						dest.getRow()+srcRowCount+ -1 + rowMultpleOffset,
						dest.getColumn()+srcColCount -1 + colMultipleOffset);
				copyCell0(src,srcBuffer,destRegion,option,rowOffset+rowMultpleOffset,columnOffset+colMultipleOffset);
			}
		}
		
	}
	

	private void copyColumnWidth(SheetRegion src, CellRegion dest) {
		// TODO zss 3.5
		
	}
	
	private void copyCell0(SheetRegion srcRegion,CellBuffer[][] srcBuffer, CellRegion destRegion,PasteOption option, int rowOffset,int columnOffset) {
		int row = destRegion.getRow();
		int col = destRegion.getColumn();
		int lastRow = destRegion.getLastRow();
		int lastColumn = destRegion.getLastColumn();
		//TODO zss 3.5 handle merge, unmerge
		for(int r = row; r <= lastRow; r++){
			for (int c = col; c <= lastColumn;c++){
				CellBuffer buffer = srcBuffer[r-row][c-col];
				if((buffer==null || buffer.getType()==CellType.BLANK) && option.isSkipBlank()){
					continue;
				}
				NCell destCell = destSheet.getCell(r,c);
				if(buffer==null){
					if(!destCell.isNull()){
						destSheet.clearCell(destCell.getRowIndex(), destCell.getColumnIndex(), destCell.getRowIndex(), destCell.getColumnIndex());
						continue;
					}
					continue;
				}
				
				switch(option.getPasteType()){
				case ALL:
					pasteValue(srcRegion,buffer,destCell,true,rowOffset,columnOffset);
					pasteStyle(buffer,destCell,true);//border,comment
					pasteComment(buffer,destCell);
					pasteValidation(buffer,destCell);
					break;
				case ALL_EXCEPT_BORDERS:
					pasteValue(srcRegion,buffer,destCell,true,rowOffset,columnOffset);
					pasteStyle(buffer,destCell,false);//border,comment
					pasteComment(buffer,destCell);
					pasteValidation(buffer,destCell);
					break;
				case COMMENTS:
					pasteComment(buffer,destCell);
					break;
				case FORMATS:
					pasteFormat(buffer,destCell);
					break;
				case FORMULAS_AND_NUMBER_FORMATS:
					pasteFormat(buffer,destCell);
				case FORMULAS:
					pasteValue(srcRegion,buffer,destCell,true,rowOffset,columnOffset);
					break;
				case VALIDATAION:
					pasteValidation(buffer,destCell);
				case VALUES_AND_NUMBER_FORMATS:
					pasteFormat(buffer,destCell);
				case VALUES:
					pasteValue(srcRegion,buffer,destCell,false,rowOffset,columnOffset);
					break;
				case COLUMN_WIDTHS:
					break;				
				}
				
			}
		}
	}

	private void pasteFormat(CellBuffer buffer, NCell destCell) {
		String srcFormat = buffer.getFormula();
		NCellStyle destStyle = destCell.getCellStyle();
		String destFormat = destStyle.getDataFormat();
		if(!destFormat.equals(srcFormat)){
			StyleUtil.setDataFormat(destSheet, destCell.getRowIndex(), destCell.getColumnIndex(), destFormat);
		}
	}

	private void pasteValidation(CellBuffer buffer, NCell destCell) {
		NDataValidation srcValidation = buffer.getValidation();
		//TODO, base on original validation data structure, it it very complicated to past a validation, consider to change the structure.
	}

	private void pasteComment(CellBuffer buffer, NCell destCell) {
		NComment comment = buffer.getComment();
		
		destCell.setComment(comment==null?null:((AbstractCommentAdv)comment).clone());
	}

	private void pasteStyle(CellBuffer buffer, NCell destCell, boolean pasteBorder) {
		if(destCell.getCellStyle()==defaultStyle && buffer.getStyle()==defaultStyle){
			return;
		}
		if(pasteBorder){
			destCell.setCellStyle(buffer.getStyle());
		}else{
			NCellStyle newStyle = book.createCellStyle(buffer.getStyle(), true);
			NCellStyle destStyle = destCell.getCellStyle();
			//keep original border
			newStyle.setBorderBottom(destStyle.getBorderBottom());
			newStyle.setBorderBottomColor(destStyle.getBorderBottomColor());
			newStyle.setBorderTop(destStyle.getBorderTop());
			newStyle.setBorderTopColor(destStyle.getBorderTopColor());
			newStyle.setBorderRight(destStyle.getBorderRight());
			newStyle.setBorderRightColor(destStyle.getBorderRightColor());
			newStyle.setBorderLeft(destStyle.getBorderLeft());
			newStyle.setBorderLeftColor(destStyle.getBorderLeftColor());
			destCell.setCellStyle(newStyle);
		}
	}

	private void pasteValue(SheetRegion sreRegion,CellBuffer buffer, NCell destCell, boolean pasteFormula,int rowOffset,int columnOffset) {
		if(pasteFormula){
			String formula = buffer.getFormula();
			if(formula!=null){
				//TODO zss 3.5 , shift non-absolute formula
				//TODO zss 3.5 rename sheetName
				FormulaEngine engine = getFormulaEignin();
				//TODO zss 3.5 should shift regardless sheet
//				FormulaExpression expr = engine.shift(formula, sreRegion.getRegion(), rowOffset, columnOffset, 
//						new FormulaParseContext(destSheet, null));
				FormulaExpression expr = engine.move(formula,sreRegion, rowOffset, columnOffset, 
						new FormulaParseContext(destSheet, null));
				System.out.println(">>> Shift "+formula+" to "+expr.getFormulaString());
				//copy paste doesn't need to handle sheet rename
				
				destCell.setFormulaValue(expr.getFormulaString());
				return;
			}
		}
		destCell.setValue(buffer.getValue());
	}

	FormulaEngine formulaEngine;
	
	private FormulaEngine getFormulaEignin() {
		if(formulaEngine == null){
			formulaEngine = EngineFactory.getInstance().createFormulaEngine();
		}
		return formulaEngine;
	}


	private CellBuffer[][] prepareBuffer(SheetRegion src, PasteOption option) {
		int srcColCount = src.getColumnCount();
		int srcRowCount = src.getRowCount();
		CellBuffer[][] srcBuffer = new CellBuffer[srcRowCount][srcColCount];
		NSheet srcSheet = src.getSheet();
		for(int r = 0; r < srcRowCount;r++){
			for(int c = 0; c < srcColCount;c++){
				NCell srcCell = srcSheet.getCell(r,c);
				if(srcCell.isNull()) //to avoid uncessary create
					continue;
				
				CellBuffer buffer = srcBuffer[r][c] = new CellBuffer(r,c);
				buffer.setType(srcCell.getType());
				
				switch(option.getPasteType()){
				case ALL:
				case ALL_EXCEPT_BORDERS:
					copyValue(buffer,srcCell,true);
					buffer.setStyle(srcCell.getCellStyle());
					buffer.setComment(srcCell.getComment());
					buffer.setValidation(srcSheet.getDataValidation(r, c));
					break;
				case COMMENTS:
					buffer.setComment(srcCell.getComment());
					break;
				case FORMATS:
					buffer.setStyle(srcCell.getCellStyle());
					break;
				case FORMULAS_AND_NUMBER_FORMATS:
					buffer.setStyle(srcCell.getCellStyle());
				case FORMULAS:
					copyValue(buffer,srcCell,true);
					break;
				case VALIDATAION:
					buffer.setValidation(srcSheet.getDataValidation(r, c));
				case VALUES_AND_NUMBER_FORMATS:
					buffer.setStyle(srcCell.getCellStyle());
				case VALUES:
					copyValue(buffer,srcCell,false);
					break;
				case COLUMN_WIDTHS:
					break;
				}
			}
		}
		return srcBuffer;
	}

	private void copyValue(CellBuffer buffer, NCell srcCell, boolean copyFormula) {
		if(copyFormula && srcCell.getType() == CellType.FORMULA){
			buffer.setFormula(srcCell.getFormulaValue());
		}else{
			buffer.setValue(srcCell.getValue());
		}
	}

}
