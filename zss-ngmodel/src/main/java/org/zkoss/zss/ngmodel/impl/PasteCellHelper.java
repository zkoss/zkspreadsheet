package org.zkoss.zss.ngmodel.impl;

import java.util.Collection;
import java.util.LinkedList;
import java.util.List;

import org.zkoss.zss.ngapi.impl.StyleUtil;
import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.InvalidateModelOpException;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NColumn;
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

/*package*/ class PasteCellHelper {

	private final NSheet destSheet;
	private final NBook book;
	private final NCellStyle defaultStyle;
	public PasteCellHelper(NSheet destSheet){
		this.destSheet = destSheet;
		this.book = destSheet.getBook();
		defaultStyle = book.getDefaultCellStyle();
	}
	
	static class CellBuffer{
		CellType type;
		Object value;
		String formula;
		NCellStyle style;
		
		NComment comment;
		NDataValidation validation;
		
		public CellBuffer(){
		}
		public CellType getType() {
			return type;
		}
		public void setType(CellType type) {
			this.type = type;
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
	
	public CellRegion pasteCell(SheetRegion src, CellRegion dest, PasteOption option) {
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
		
		int rowOffset = dest.getRow() - src.getRow();
		int columnOffset = dest.getColumn() - src.getColumn();
	
		int destColCount = dest.getColumnCount();
		int destRowCount = dest.getRowCount();
		
		if(option.getPasteType()==PasteType.COLUMN_WIDTHS){
			if(option.isCut()){
				throw new InvalidateModelOpException("can't do cut when copying column width");
			}else if(option.isTranspose()){
				throw new InvalidateModelOpException("can't do transport when copying column width");
			}
			int[] widthBuffer = prepareColumnWidth(src);
			int srcColCount = widthBuffer.length;
			boolean wrongColMultiple = (destColCount>1 && destColCount%srcColCount!=0);
			int colMultiple = destColCount<=1||wrongColMultiple?1:destColCount/srcColCount;
			for(int j=0;j<colMultiple;j++){
					int colMultipleOffset = j*srcColCount;
					CellRegion destRegion = new CellRegion(dest.getRow(),dest.getColumn()+colMultipleOffset,
							dest.getRow(),dest.getColumn()+srcColCount -1 + colMultipleOffset);
					pasteColumnWidth(widthBuffer,destRegion);
			}
			return new CellRegion(dest.getRow(),dest.getColumn(),dest.getRow(),dest.getColumn()+srcColCount*colMultiple-1);
		}
		PasteType pasteType = option.getPasteType();
		boolean handleMerge = shouldHandleMerge(pasteType);
		
		CellRegion srcRegion = src.getRegion();
		if(handleMerge){
			Collection<CellRegion> srcOverlaps = srcSheet.getOverlapsMergedRegions(srcRegion,true);
			if(srcOverlaps.size()>0){
				throw new InvalidateModelOpException("Can't copy "+srcRegion.getReferenceString()+" which overlaps merge area "+srcOverlaps.iterator().next().getReferenceString());
			}
		}
		
		//the buffer might be transported
		CellBuffer[][] srcBuffer = prepareCellBuffer(src,option);
		Collection<CellRegion> mergeBuffer = null;
		if(handleMerge){
			mergeBuffer = prepareMergeRegionBuffer(src,option);
		}
		
		int srcColCount = srcBuffer[0].length;
		int srcRowCount = srcBuffer.length;
		
		boolean wrongRowMultiple = (destRowCount>1 && destRowCount%srcRowCount!=0);
		boolean wrongColMultiple = (destColCount>1 && destColCount%srcColCount!=0);

		
		boolean wrongMultiple = wrongRowMultiple||wrongColMultiple;
		int rowMultiple = destRowCount<=1||wrongMultiple?1:destRowCount/srcRowCount;
		int colMultiple = destColCount<=1||wrongMultiple?1:destColCount/srcColCount;

		if(option.isCut()){
			//clear the src's value and merge
			clearCell(src);
			clearMergeRegion(src);
		}

		for(int i=0;i<rowMultiple;i++){
			for(int j=0;j<colMultiple;j++){
				int rowMultpleOffset = i*srcRowCount;
				int colMultipleOffset = j*srcColCount;
				CellRegion destRegion = new CellRegion(dest.getRow()+rowMultpleOffset,dest.getColumn()+colMultipleOffset,
						dest.getRow()+srcRowCount+ -1 + rowMultpleOffset,
						dest.getColumn()+srcColCount -1 + colMultipleOffset);
				pasteCell(srcBuffer,destRegion,option,rowOffset+rowMultpleOffset,columnOffset+colMultipleOffset);
				
				if(mergeBuffer!=null && mergeBuffer.size()>0){
					pasteMergeRegion(mergeBuffer,rowOffset+rowMultpleOffset,columnOffset+colMultipleOffset);
				}
			}
		}
		
		return new CellRegion(dest.getRow(),dest.getColumn(),
				dest.getRow()+srcRowCount*rowMultiple-1,
				dest.getColumn()+srcColCount*colMultiple-1);
	}
	

	FormulaEngine formulaEngine;
	private FormulaEngine getFormulaEignin() {
		if(formulaEngine == null){
			formulaEngine = EngineFactory.getInstance().createFormulaEngine();
		}
		return formulaEngine;
	}


	private boolean shouldHandleMerge(PasteType pasteType) {
		return pasteType == PasteType.ALL ||
				pasteType == PasteType.ALL_EXCEPT_BORDERS ||
				pasteType == PasteType.FORMATS;
	}


	private int[] prepareColumnWidth(SheetRegion src){
		int column = src.getColumn();
		int lastColumn = src.getLastColumn();
		int srcColCount = src.getColumnCount();
		int[] widthBuffer = new int[srcColCount];
		NSheet srcSheet = src.getSheet();
		for(int c = column; c <= lastColumn;c++){
			widthBuffer[c-column] = srcSheet.getColumn(c).getWidth();
		}
		return widthBuffer;
	}


	private List<CellRegion> prepareMergeRegionBuffer(SheetRegion src,
			PasteOption option) {
		List<CellRegion> mergeBuffer = new LinkedList<CellRegion>();
		boolean transpose = option.isTranspose();
		for(CellRegion region:src.getSheet().getContainsMergedRegions(src.getRegion())){
			if(transpose){
				int rowAnchor = src.getRow();
				int columnAnchor = src.getColumn();
				int r = rowAnchor + region.getColumn()-columnAnchor;
				int c = columnAnchor + region.getRow()-rowAnchor;
				int lr = r + region.getColumnCount()-1;
				int lc = c + region.getRowCount()-1;
				region = new CellRegion(r, c, lr, lc);
			}
			mergeBuffer.add(region);
		}
		return mergeBuffer;
	}


	private CellBuffer[][] prepareCellBuffer(SheetRegion src, PasteOption option) {
		int row = src.getRow();
		int column = src.getColumn();
		int lastRow = src.getLastRow();
		int lastColumn = src.getLastColumn();
		
		boolean transpose = option.isTranspose();
		int srcRowCount = transpose?src.getColumnCount():src.getRowCount();
		int srcColCount = transpose?src.getRowCount():src.getColumnCount();
		
		
		CellBuffer[][] srcBuffer = new CellBuffer[srcRowCount][srcColCount];
		NSheet srcSheet = src.getSheet();
		for(int r = row; r <= lastRow;r++){
			for(int c = column; c <= lastColumn;c++){
				NCell srcCell = srcSheet.getCell(r,c);
				if(srcCell.isNull()) //to avoid unnecessary create
					continue;
				
				CellBuffer buffer = srcBuffer[transpose?c-column:r-row][transpose?r-row:c-column] = new CellBuffer();
				
				buffer.setType(srcCell.getType());
				
				switch(option.getPasteType()){
				case ALL:
				case ALL_EXCEPT_BORDERS:
					prepareValue(buffer,srcCell,true);
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
					prepareValue(buffer,srcCell,true);
					break;
				case VALIDATAION:
					buffer.setValidation(srcSheet.getDataValidation(r, c));
				case VALUES_AND_NUMBER_FORMATS:
					buffer.setStyle(srcCell.getCellStyle());
				case VALUES:
					prepareValue(buffer,srcCell,false);
					break;
				case COLUMN_WIDTHS:
					break;
				}
			}
		}
		return srcBuffer;
	}


	private void prepareValue(CellBuffer buffer, NCell srcCell, boolean copyFormula) {
		if(copyFormula && srcCell.getType() == CellType.FORMULA){
			buffer.setFormula(srcCell.getFormulaValue());
		}else{
			buffer.setValue(srcCell.getValue());
		}
	}


	private void clearMergeRegion(SheetRegion src) {
		src.getSheet().removeMergedRegion(src.getRegion(), true);
		
//		for(CellRegion merge:getOverlapsMergedRegions(src.getRegion(),false)){
//			src.getSheet().removeMergedRegion(merge);
//		}
	}


	private void clearCell(SheetRegion src) {
		int row = src.getRow();
		int column = src.getColumn();
		int lastRow = src.getLastRow();
		int lastColumn = src.getLastColumn();
		NSheet srcSheet = src.getSheet();
		for(int r = row ; r<=lastRow;r++){
			for(int c= column; c<=lastColumn;c++){
				NCell cell = srcSheet.getCell(r,c);
				if(!cell.isNull()){
					srcSheet.clearCell(new CellRegion(r,c));
				}
			}
		}
	}


	private void pasteMergeRegion(Collection<CellRegion> mergeBuffer, int rowOffset,int columnOffset) {
		for(CellRegion merge:mergeBuffer){
			CellRegion newMerge = new CellRegion(merge.getRow()+rowOffset,merge.getColumn()+columnOffset,
					merge.getLastRow()+rowOffset,merge.getLastColumn()+columnOffset);
			
//			for(CellRegion overlap:destSheet.getOverlapsMergedRegions(newMerge,false)){
//				//unmerge the overlaps
//				destSheet.removeMergedRegion(overlap);
//			}
			destSheet.removeMergedRegion(newMerge, true);
			
			destSheet.addMergedRegion(newMerge);
		}
	}


	private void pasteColumnWidth(int[] widthBuffer, CellRegion dest) {
		int col = dest.getColumn();
		int lastColumn = dest.getLastColumn();
		for (int c = col; c <= lastColumn;c++){
			destSheet.getColumn(c).setWidth(widthBuffer[c-col]);
		}
	}
	
	private void pasteCell(CellBuffer[][] srcBuffer, CellRegion destRegion,PasteOption option, int rowOffset,int columnOffset) {
		int row = destRegion.getRow();
		int col = destRegion.getColumn();
		int lastRow = destRegion.getLastRow();
		int lastColumn = destRegion.getLastColumn();
		boolean transport = option.isTranspose();
		for(int r = row; r <= lastRow; r++){
			for (int c = col; c <= lastColumn;c++){
				CellBuffer buffer = srcBuffer[r-row][c-col];
				if((buffer==null || buffer.getType()==CellType.BLANK) && option.isSkipBlank()){
					continue;
				}
				
				//unmerge region if it is overlaps and not at first cell
				CellRegion region = destSheet.getMergedRegion(r, c);
				if(region!=null && (region.getRow()!=r || region.getColumn()!=c)){
					destSheet.removeMergedRegion(region, true);
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
					pasteValue(buffer,destCell,true,rowOffset,columnOffset,option);
					pasteStyle(buffer,destCell,true);//border,comment
					pasteComment(buffer,destCell);
					pasteValidation(buffer,destCell);
					break;
				case ALL_EXCEPT_BORDERS:
					pasteValue(buffer,destCell,true,rowOffset,columnOffset,option);
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
					pasteValue(buffer,destCell,true,rowOffset,columnOffset,option);
					break;
				case VALIDATAION:
					pasteValidation(buffer,destCell);
				case VALUES_AND_NUMBER_FORMATS:
					pasteFormat(buffer,destCell);
				case VALUES:
					pasteValue(buffer,destCell,false,rowOffset,columnOffset,option);
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

	private void pasteValue(CellBuffer buffer, NCell destCell, boolean pasteFormula,int rowOffset,int columnOffset,PasteOption option) {
		if(pasteFormula){
			String formula = buffer.getFormula();
			if(formula!=null){
				boolean transpose = option.isTranspose();
				//TODO zss 3.5 , shift non-absolute formula
				//TODO zss 3.5 rename sheetName
				FormulaEngine engine = getFormulaEignin();
				//TODO zss 3.5 should shift regardless sheet
//				FormulaExpression expr = engine.shift(formula,rowOffset, columnOffset, 
//						new FormulaParseContext(destSheet, null));
				//TODO this a temporary
				FormulaExpression expr = engine.move(formula, new SheetRegion(destSheet,0,0,book.getMaxRowIndex(),book.getMaxColumnIndex()), 
						transpose?columnOffset:rowOffset, transpose?rowOffset:columnOffset, 
						new FormulaParseContext(destSheet, null));
				System.out.println(">>> Shift "+formula+" to "+expr.getFormulaString());
				//copy paste doesn't need to handle sheet rename
				
				destCell.setFormulaValue(expr.getFormulaString());
				return;
			}
		}
		destCell.setValue(buffer.getValue());
	}

}
