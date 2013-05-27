package org.zkoss.zss.api;

import java.util.LinkedHashMap;

import org.zkoss.zss.api.Range.ApplyBorderType;
import org.zkoss.zss.api.Range.DeleteShift;
import org.zkoss.zss.api.Range.InsertCopyOrigin;
import org.zkoss.zss.api.Range.InsertShift;
import org.zkoss.zss.api.Range.PasteOperation;
import org.zkoss.zss.api.Range.PasteType;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.CellStyle.Alignment;
import org.zkoss.zss.api.model.CellStyle.BorderType;
import org.zkoss.zss.api.model.CellStyle.FillPattern;
import org.zkoss.zss.api.model.CellStyle.VerticalAlignment;
import org.zkoss.zss.api.model.Color;
import org.zkoss.zss.api.model.Font;
import org.zkoss.zss.api.model.Font.Boldweight;
import org.zkoss.zss.api.model.Font.Underline;
import org.zkoss.zss.api.model.Sheet;

/**
 * The utility to help UI to deal with user's cell operation of a Range.
 * This utility is the default implementation for handling a spreadsheet, it is also the example for calling {@link Range} APIs 
 * @author dennis
 * @since 3.0.0
 */
public class CellOperationUtil {

	
	static private class Result<T> {
		T r;
		private Result(){}
		private Result(T r){
			this.r = r;
		}
		
		public T get(){
			return r;
		}
		
		public void set(T r){
			this.r = r;
		}
	}
	
	/**
	 * Cuts data and style from src to destination
	 * @param src source range
	 * @param dest destination range
	 * @return false if sheet is protected
	 */
	public static boolean cut(Range src, final Range dest) {
		final Result<Boolean> result = new Result<Boolean>();
		if(src.isProtected()){
			return false;
		}
		if(dest.isProtected()){
			return false;
		}
		//use batch-runner to run multiple range operation
		src.sync(new RangeRunner() {
			public void run(Range range) {
				boolean r = range.paste(dest);
				if(r){
					range.clearContents();// it removes value and formula only
					range.clearStyles();
				}
				result.set(r);
			}
		});

		return result.get();
	}

	/**
	 * Paste data and style from src to destination
	 * @param src source range
	 * @param dest destination range
	 * @return false if sheet is protected
	 */
	public static boolean paste(Range src, Range dest) {
		if(dest.isProtected()){
			return false;
		}
		return src.paste(dest);
	}
	
	/**
	 * Paste formula only from src to destination
	 * @param src source range
	 * @param dest destination range
	 * @return false if sheet is protected
	 */
	public static boolean pasteFormula(Range src, Range dest) {
		return pasteSpecial(src, dest, PasteType.FORMULAS, PasteOperation.NONE, false, false);
	}
	/**
	 * Paste value only from src to destination
	 * @param src source range
	 * @param dest destination range
	 * @return false if sheet is protected
	 */
	public static boolean pasteValue(Range src, Range dest) {
		return pasteSpecial(src, dest, PasteType.VALUES, PasteOperation.NONE, false, false);
	}
	/**
	 * Paste all (except border) from src to destination
	 * @param src source range
	 * @param dest destination range
	 * @return false if sheet is protected
	 */
	public static boolean pasteAllExceptBorder(Range src, Range dest) {
		return pasteSpecial(src, dest, PasteType.ALL_EXCEPT_BORDERS, PasteOperation.NONE, false, false);
	}
	
	/**
	 * Paste and transpose from src to destination
	 * @param src source range
	 * @param dest destination range
	 * @return false if sheet is protected
	 */
	public static boolean pasteTranspose(Range src, Range dest) {
		return pasteSpecial(src, dest, PasteType.ALL, PasteOperation.NONE, false, true);
	}
	
	/**
	 * Paste according the argument from src to destination
	 * @param src source range
	 * @param dest destination range
	 * @param pasteType paste type
	 * @param pasteOperation paste operation
	 * @param skipBlank skip blank
	 * @param transpose transpose
	 * @return false if sheet is protected
	 */
	public static boolean pasteSpecial(Range src, Range dest,PasteType pasteType, PasteOperation pasteOperation, boolean skipBlank, boolean transpose){
		if(dest.isProtected()){
			return false;
		}
		return src.pasteSpecial(dest, pasteType, pasteOperation, skipBlank, transpose);
	}

	/**
	 * Apply font to cells in the range
	 * @param range range to be applied
	 * @param fontName the font name
	 */
	public static void applyFontName(Range range,final String fontName){
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
				//the font name(family) is equals, not need to set it.
				return oldFont.getFontName().equals(fontName);
			}
			
			public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
				//use the new font name to search it.
				return cellRange.getCellStyleHelper().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), fontName, 
						oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(Range range,CellStyle cellstyle, Font newfont) {
				newfont.setFontName(fontName);
			}
		});
	}

	/**
	 * Apply font height to cells in the range, it will also enlarge the row height if row height is smaller than font height 
	 * @param range range to be applied
	 * @param fontHeight the font height in twpi (1/20 point)
	 */
	public static void applyFontHeight(Range range, final short fontHeight) {
		//fontHeight = twip
		final Sheet sheet = range.getSheet(); 
		final LinkedHashMap<Integer,Integer> newRowSize = new LinkedHashMap<Integer, Integer>();
		
		if(range.isProtected())
			return;
		
		final int fpx = UnitUtil.twipToPx(fontHeight);
		
		range.sync(new RangeRunner(){
			@Override
			public void run(Range range) {
				applyFontStyle(range, new FontStyleApplier() {
					public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
						//the font name(family) is equals, not need to set it.
						return oldFont.getFontHeight() == fontHeight;
					}
					
					public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
						int px = sheet.getRowHeight(cellRange.getRow());//rowHeight in px
						
						if(fpx>px){
							Integer maxpx = newRowSize.get(cellRange.getRow());
							if(maxpx==null || maxpx.intValue()<fpx){
								newRowSize.put(cellRange.getRow(), fpx);
							}
						}
						
						//use the new font name to search it.
						return cellRange.getCellStyleHelper().findFont(oldFont.getBoldweight(), oldFont.getColor(), fontHeight, oldFont.getFontName(), 
								oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
					}
					
					public void apply(Range cellRange,CellStyle cellstyle, Font newfont) {
						newfont.setFontHeight(fontHeight);
					}
				});
				//enhancement, also adjust row height.
				for(Integer row : newRowSize.keySet()){
					Ranges.range(sheet,row.intValue(),0).setRowHeight(newRowSize.get(row).intValue()+4);//4 is padding
				}
			}
		});
	}
	
	/**
	 * Apply the font size to cells in the range, it will also enlarge the row height if row height is smaller than font size 
	 * @param range range to be applied
	 * @param point the font size in point
	 */
	public static void applyFontSize(Range range, final short point) {
		//fontSize = fontHeightInPoints = fontHeight/20
		applyFontHeight(range,(short)(point*20));
	}
	
	/**
	 * Interface for helping apply font style.
	 * @author dennis
	 */
	public interface FontStyleApplier {
		/** Should ignore this cellRange**/
		public boolean ignore(Range cellRange,CellStyle oldCellstyle,Font oldFont);
		/** Find the font to apply to new style, return null mean not found, and will create a new font**/
		public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont);
		/** apply style to new font, will be call when {@code #search() return null}**/
		public void apply(Range cellRange,CellStyle newCellstyle,Font newfont);
	}
	
	/**
	 * Apply font style according to the {@link FontStyleApplier}
	 * @param range the range to be applied
	 * @param applyer the font style appliler
	 */
	public static void applyFontStyle(Range range,final FontStyleApplier applyer){
		//use use cell visitor to visit cells respectively
		//use BOOK level lock because we will create BOOK level Object (style, font)
		if(range.isProtected())
			return;
		range.visit(new CellVisitor(){
			@Override
			public boolean visit(Range cellRange) {
				CellStyle ostyle = cellRange.getCellStyle();
				Font ofont = ostyle.getFont();
				//ignore or not, to prevent unnecessary creation
				if(applyer.ignore(cellRange,ostyle,ofont)){
					return true;//continue visit
				}
				//2.create a new style from old one, the original one is shared between cells
				//TODO what if it is the last referenced, could I just use it? who will delete old if I use new
				CellStyle nstyle = cellRange.getCellStyleHelper().createCellStyle(ostyle);
				
				Font nfont = null;
				//3.search if there any font already in book
				nfont = applyer.search(cellRange,ostyle,ofont);
				if( nfont== null){
					//create new font base old font if not found
					nfont = cellRange.getCellStyleHelper().createFont(ofont);
					nstyle.setFont(nfont);
					
					//set the apply
					applyer.apply(cellRange,nstyle, nfont);
				}else{
					nstyle.setFont(nfont);
				}
				
				
				cellRange.setCellStyle(nstyle);
				return true;
			}

			@Override
			public boolean ignoreIfNotExist(int row, int column) {
				return false;
			}

			@Override
			public boolean createIfNotExist(int row, int column) {
				return true;
			}});
	}

	
	/**
	 * Apply font bold-weight to cells in the range
	 * @param range the range to be applied
	 * @param boldweight the font bold-weight
	 */
	public static void applyFontBoldweight(Range range,final Boldweight boldweight) {
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
				return oldFont.getBoldweight().equals(boldweight);
			}
			
			public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
				return cellRange.getCellStyleHelper().findFont(boldweight, oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), 
						oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(Range range,CellStyle newCellstyle, Font newfont) {
				newfont.setBoldweight(boldweight);
			}
		});
	}

	/**
	 * Apply font italic to cells in the range
	 * @param range the range to be applied
	 * @param italic the font italic
	 */
	public static void applyFontItalic(Range range, final boolean italic) {
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
				return oldFont.isItalic()==italic;
			}
			
			public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
				return cellRange.getCellStyleHelper().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), 
						italic, oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(Range range,CellStyle newCellstyle, Font newfont) {
				newfont.setItalic(italic);
			}
		});
	}

	/**
	 * Apply font strike-out to cells in the range 
	 * @param range the range to be applied
	 * @param strikeout font strike-out
	 */
	public static void applyFontStrikeout(Range range, final boolean strikeout) {
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
				return oldFont.isStrikeout()==strikeout;
			}
			
			public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
				return cellRange.getCellStyleHelper().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), 
						oldFont.isItalic(), strikeout, oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(Range range,CellStyle newCellstyle, Font newfont) {
				newfont.setStrikeout(strikeout);
			}
		});
	}
	
	/**
	 * Apply font underline to cells in the range
	 * @param range the range to be applied
	 * @param underline font underline
	 */
	public static void applyFontUnderline(Range range,final Underline underline) {
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
				return oldFont.getUnderline().equals(underline);
			}
			
			public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
				return cellRange.getCellStyleHelper().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), 
						oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), underline);
			}
			
			public void apply(Range range,CellStyle newCellstyle, Font newfont) {
				newfont.setUnderline(underline);
			}
		});
	}
	
	/**
	 * Apply font color to cells in the range
	 * @param range the range to be applied.
	 * @param htmlColor the color by html color syntax(#rgb-hex-code, e.x #FF00FF) 
	 */
	public static void applyFontColor(Range range, final String htmlColor) {
		final Color color = range.getCellStyleHelper().createColorFromHtmlColor(htmlColor);
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
				return oldFont.getColor().getHtmlColor().equals(htmlColor);
			}
			
			public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
				return cellRange.getCellStyleHelper().findFont(oldFont.getBoldweight(), color, oldFont.getFontHeight(), oldFont.getFontName(), 
						oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(Range range,CellStyle newCellstyle, Font newfont) {
				
				//check it in XSSFCellStyle , it is just a delegator, but do some thing in HSSF
				newfont.setColor(color);
				
				//TODO need to check with Henri, and Sam, why current implementation doesn't call set font color directly
				//TODO call style's set font color will cause set color after set a theme color issue(after clone form default)
//				newCellstyle.setFontColor(color); 
			}
		});
	}
	
	/**
	 * Apply backgound-color to cells in the range
	 * @param range the range to be applied
	 * @param htmlColor the color by html color syntax(#rgb-hex-code, e.x #FF00FF) 
	 */
	public static void applyBackgroundColor(Range range, final String htmlColor) {
		final Color color = range.getCellStyleHelper().createColorFromHtmlColor(htmlColor);
		applyCellStyle(range, new CellStyleApplier() {

			public boolean ignore(Range cellRange, CellStyle oldCellstyle) {
				Color ocolor = oldCellstyle.getBackgroundColor();
				return ocolor.equals(color);
			}

			public void apply(Range cellRange, CellStyle newCellstyle) {
				newCellstyle.setBackgroundColor(color);
				
				FillPattern patternType = newCellstyle.getFillPattern();
				if (patternType == FillPattern.NO_FILL) {
					newCellstyle.setFillPattern(FillPattern.SOLID_FOREGROUND);
				}
			}
		});
	}
	
	/**
	 * Apply data-format to cells in the range
	 * @param range the range to be applied
	 * @param format the data format
	 */
	public static void applyDataFormat(Range range, final String format) {
		applyCellStyle(range, new CellStyleApplier() {

			public boolean ignore(Range cellRange, CellStyle oldCellstyle) {
				String oformat = oldCellstyle.getDataFormat();
				return oformat.equals(format);
			}

			public void apply(Range cellRange, CellStyle newCellstyle) {
				newCellstyle.setDataFormat(format);
			}
		});
	}

	/**
	 * Apply alignment to cells in the range
	 * @param range the range to be applied
	 * @param alignment the alignement
	 */
	public static void applyAlignment(Range range,final Alignment alignment) {
		applyCellStyle(range, new CellStyleApplier() {

			public boolean ignore(Range cellRange, CellStyle oldCellstyle) {
				Alignment oldalign = oldCellstyle.getAlignment();
				return oldalign.equals(alignment);
			}

			public void apply(Range cellRange, CellStyle newCellstyle) {
				newCellstyle.setAlignment(alignment);
			}
		});
	}

	/**
	 * Apply vertical-alignment to cells in the range
	 * @param range the range to be applied
	 * @param alignment vertical alignment
	 */
	public static void applyVerticalAlignment(Range range,final VerticalAlignment alignment) {
		applyCellStyle(range, new CellStyleApplier() {

			public boolean ignore(Range cellRange, CellStyle oldCellstyle) {
				VerticalAlignment oldalign = oldCellstyle.getVerticalAlignment();
				return oldalign.equals(alignment);
			}

			public void apply(Range cellRange, CellStyle newCellstyle) {
				newCellstyle.setVerticalAlignment(alignment);
			}
		});
	}
	
	/**
	 * Interface for help apply cell style
	 */
	public interface CellStyleApplier {
		/** should ignore this cellRange**/
		public boolean ignore(Range cellRange,CellStyle oldCellstyle);
		/** apply style to new cell**/
		public void apply(Range cellRange,CellStyle newCellstyle);
	}
	
	/**
	 * Apply style according to the cell style applier
	 * @param range the range to be applied
	 * @param applyer
	 */
	public static void applyCellStyle(Range range,
			final CellStyleApplier applyer) {
		// use use cell visitor to visit cells respectively
		// use BOOK level lock because we will create BOOK level Object (style,
		// font)
		if(range.isProtected())
			return;
		range.visit(new CellVisitor() {
			@Override
			public boolean visit(Range cellRange) {
				CellStyle ostyle = cellRange.getCellStyle();
				
				// ignore or not, to prevent unnecessary creation
				if (applyer.ignore(cellRange, ostyle)) {
					return true;//continue visit
				}
				// 2.create a new style from old one, the original one is shared
				// between cells
				// TODO what if it is the last referenced, could I just use it?
				// who will delete old if I use new
				CellStyle nstyle = cellRange.getCellStyleHelper().createCellStyle(
						ostyle);

				// set the apply
				applyer.apply(cellRange, nstyle);

				cellRange.setCellStyle(nstyle);
				return true;
			}

			@Override
			public boolean ignoreIfNotExist(int row, int column) {
				return false;
			}

			@Override
			public boolean createIfNotExist(int row, int column) {
				return true;
			}
		});
	}
	
	/**
	 * Apply border to cells in the range
	 * @param range the range to be applied
	 * @param type the apply type
	 * @param borderType the border type
	 * @param htmlColor the color of border(#rgb-hex-code, e.x #FF00FF) 
	 */
	public static void applyBorder(Range range,ApplyBorderType type,BorderType borderType,String htmlColor){
		if(range.isProtected())
			return;
		//use range api directly,
		range.applyBorders(type, borderType, htmlColor);
	}
	
	
	/**
	 * Toggle merge/unMerge of the range, if merging it will also set alignment to center
     * @param range the range to be applied
	 */
	public static void toggleMergeCenter(Range range){
		if(range.isProtected())
			return;
		range.sync(new RangeRunner() {
			public void run(Range range) {
				if(range.hasMergedCell()){
					range.unMerge();
				}else{
					range.merge(false);
					//align the left/top one
					applyAlignment(range.toCellRange(0,0),Alignment.CENTER);
				}
			}
		});
	}
	
	/**
	 * merge the range  
	 * @param range the range to be merge
	 * @param across true if merge horizontally
	 */
	public static void merge(Range range,boolean across){
		if(range.isProtected())
			return;
		
		range.merge(across);
	}
	
	/**
	 * Un-merge the range
	 * @param range the range to be un-merge
	 */
	public static void unMerge(Range range){
		if(range.isProtected())
			return;
		range.unMerge();
	}
	
	/**
	 * Apply text-warp to cells in the range
     * @param range the range to be applied
	 * @param wraptext wrap text or not
	 */
	public static void applyWrapText(Range range,final boolean wraptext) {
		applyCellStyle(range, new CellStyleApplier() {

			public boolean ignore(Range cellRange, CellStyle oldCellstyle) {
				boolean oldwrap = oldCellstyle.isWrapText();
				return oldwrap==wraptext;
			}

			public void apply(Range cellRange, CellStyle newCellstyle) {
				newCellstyle.setWrapText(wraptext);
			}
		});
	}

	/**
	 * Clear contents
	 * @param range the range to be cleared.
	 */
	public static void clearContents(Range range) {
		if(range.isProtected())
			return;
		range.clearContents();
	}
	/**
	 * Clear style
	 * @param range the range to be cleared
	 */
	public static void clearStyles(Range range) {
		if(range.isProtected())
			return;
		range.clearStyles();
	}
	
	/**
	 * Clear all 
	 * @param range the range to be cleared
	 */
	public static void clearAll(Range range){
		if(range.isProtected())
			return;

		//use batch-runner to run multiple range operation
		range.sync(new RangeRunner() {
			public void run(Range range) {
				range.clearContents();// it removes value and formula only
				range.clearStyles();
				//TODO clear hyperlink
			}
		});
	}

	/**
	 * Insert cells to the range. To insert a row, you have to call {@link Range#toRowRange()} first, to insert a column, you have to call {@link Range#toColumnRange()} first. 
	 * @param range the range to insert new cells
	 * @param shift the shift direction of original cells
	 * @param copyOrigin copy the format from nearby cells when inserting new cells 
	 */
	public static void insert(Range range, InsertShift shift,
			InsertCopyOrigin copyOrigin) {
		if(range.isProtected())
			return;
		range.insert(shift, copyOrigin);
	}
	
	/**
	 * Delete cells of the range. To delete a row, you have to call {@link Range#toRowRange()} first, to delete a column, you have to call {@link Range#toColumnRange()} first.
	 * @param range the range to delete
	 * @param shift the shift direction when deleting.
	 */
	public static void delete(Range range, DeleteShift shift) {
		if(range.isProtected())
			return;
		range.delete(shift);
	}

	/**
	 * Sort range
	 * @param range the range to sort
	 * @param desc true for descent, false for ascent
	 */
	public static void sort(Range range, boolean desc) {
		if(range.isProtected())
			return;
		range.sort(desc);
	}

	/**
	 * Hide the range. To hide a row, you have to call {@link Range#toRowRange()} first, to hide a column, you have to call {@link Range#toColumnRange()}
	 * @param range the range to hide
	 */
	public static void hide(Range range) {
		if (range.isProtected())
			return;
		range.setHidden(true);
	}
	/**
	 * Unhide the range. To unhide a row, you have to call {@link Range#toRowRange()} first, to unhide a column, you have to call {@link Range#toColumnRange()}
	 * @param range the range to un-hide
	 */
	public static void unHide(Range range) {
		if (range.isProtected())
			return;
		range.setHidden(false);
	}
	
	/**
	 * Shifts/moves cells with a offset row and column
	 * @param range the range to shift
	 * @param rowOffset the row offset
	 * @param colOffset the column offset
	 */
	public static void shift(Range range, int rowOffset, int colOffset) {
		if(range.isProtected())
			return;
		range.shift(rowOffset, colOffset);
	}
}
