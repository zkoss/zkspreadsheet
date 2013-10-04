/* CellOperationUtil.java

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

import org.zkoss.zss.api.Range.ApplyBorderType;
import org.zkoss.zss.api.Range.AutoFillType;
import org.zkoss.zss.api.Range.DeleteShift;
import org.zkoss.zss.api.Range.InsertCopyOrigin;
import org.zkoss.zss.api.Range.InsertShift;
import org.zkoss.zss.api.Range.PasteOperation;
import org.zkoss.zss.api.Range.PasteType;
import org.zkoss.zss.api.Range.SortDataOption;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.CellStyle.Alignment;
import org.zkoss.zss.api.model.CellStyle.BorderType;
import org.zkoss.zss.api.model.CellStyle.FillPattern;
import org.zkoss.zss.api.model.CellStyle.VerticalAlignment;
import org.zkoss.zss.api.model.Color;
import org.zkoss.zss.api.model.EditableCellStyle;
import org.zkoss.zss.api.model.EditableFont;
import org.zkoss.zss.api.model.Font;
import org.zkoss.zss.api.model.Font.Boldweight;
import org.zkoss.zss.api.model.Font.Underline;
import org.zkoss.zss.api.model.Hyperlink.HyperlinkType;

/**
 * The utility to help UI to deal with user's cell operation of a {@link Range}.
 * This utility is the default implementation for handling user operations for cells, it is also the example for calling {@link Range} APIs 
 * @author dennis
 * @since 3.0.0
 */
public class CellOperationUtil {
	
	/**
	 * Cuts data and style from src to destination
	 * @param src source range
	 * @param dest destination range
	 * @return a Range contains the final pasted range. paste to a protected sheet will always cause paste return null.
	 */
	public static Range cut(Range src, final Range dest) {
		if(src.isProtected()){
			return null;
		}
		if(dest.isProtected()){
			return null;
		}
		return src.paste(dest, true);
	}
	
	

	/**
	 * Paste data and style from src to destination
	 * @param src source range
	 * @param dest destination range
	 * @return a Range contains the final pasted range. paste to a protected sheet will always cause paste return null.
	 */
	public static Range paste(Range src, Range dest) {
		if(dest.isProtected()){
			return null;
		}
		return src.paste(dest);
	}
	
	/**
	 * Paste formula only from src to destination
	 * @param src source range
	 * @param dest destination range
	 * @return a Range contains the final pasted range. paste to a protected sheet will always cause paste return null.
	 */
	public static Range pasteFormula(Range src, Range dest) {
		return pasteSpecial(src, dest, PasteType.FORMULAS, PasteOperation.NONE, false, false);
	}
	/**
	 * Paste value only from src to destination
	 * @param src source range
	 * @param dest destination range
	 * @return a Range contains the final pasted range. paste to a protected sheet will always cause paste return null.
	 */
	public static Range pasteValue(Range src, Range dest) {
		return pasteSpecial(src, dest, PasteType.VALUES, PasteOperation.NONE, false, false);
	}
	/**
	 * Paste all (except border) from src to destination
	 * @param src source range
	 * @param dest destination range
	 * @return a Range contains the final pasted range. paste to a protected sheet will always cause paste return null.
	 */
	public static Range pasteAllExceptBorder(Range src, Range dest) {
		return pasteSpecial(src, dest, PasteType.ALL_EXCEPT_BORDERS, PasteOperation.NONE, false, false);
	}
	
	/**
	 * Paste and transpose from src to destination
	 * @param src source range
	 * @param dest destination range
	 * @return a Range contains the final pasted range. paste to a protected sheet will always cause paste return null.
	 */
	public static Range pasteTranspose(Range src, Range dest) {
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
	 *@return a Range contains the final pasted range. paste to a protected sheet will always cause paste return null.
	 */
	public static Range pasteSpecial(Range src, Range dest,PasteType pasteType, PasteOperation pasteOperation, boolean skipBlank, boolean transpose){
		if(dest.isProtected()){
			return null;
		}
		return src.pasteSpecial(dest, pasteType, pasteOperation, skipBlank, transpose);
	}

	public static FontStyleApplier getFontNameApplier(final String fontName){
		return new FontStyleApplier() {
			public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
				//the font name(family) is equals, not need to set it.
				return oldFont.getFontName().equals(fontName);
			}
			
			public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
				//use the new font name to search it.
				return cellRange.getCellStyleHelper().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), fontName, 
						oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(Range range,EditableCellStyle cellstyle, EditableFont newfont) {
				newfont.setFontName(fontName);
			}
		};
	}
	/**
	 * Apply font to cells in the range
	 * @param range range to be applied
	 * @param fontName the font name
	 */
	public static void applyFontName(Range range,final String fontName){
		applyFontStyle(range, getFontNameApplier(fontName));
	}
	
	public static FontStyleApplier getFontHeightApplier(final short fontHeight) {
		//fontHeight = twip
		final int fpx = UnitUtil.twipToPx(fontHeight);
		return new FontStyleApplier() {
			public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
				//the font name(family) is equals, not need to set it.
				return oldFont.getFontHeight() == fontHeight;
			}
			
			public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
				//use the new font name to search it.
				Font f = cellRange.getCellStyleHelper().findFont(oldFont.getBoldweight(), oldFont.getColor(), fontHeight, oldFont.getFontName(), 
						oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
				if(f!=null){
					//if font was found, the apply will be skip, so i enlarge here
					int px = cellRange.getSheet().getRowHeight(cellRange.getRow());//rowHeight in px
					if(fpx>px){
						cellRange.setRowHeight(fpx+4);//4 is padding
					}
				}
				return f;
			}
			
			public void apply(Range cellRange,EditableCellStyle cellstyle, EditableFont newfont) {
				int px = cellRange.getSheet().getRowHeight(cellRange.getRow());//rowHeight in px
				newfont.setFontHeight(fontHeight);
				if(fpx>px){//enlarge the row height
					cellRange.setRowHeight(fpx+4);//4 is padding
				}
			}
		};
	}

	/**
	 * Apply font height to cells in the range, it will also enlarge the row height if row height is smaller than font height 
	 * @param range range to be applied
	 * @param fontHeight the font height in twpi (1/20 point)
	 */
	public static void applyFontHeight(Range range, final short fontHeight) {
		applyFontStyle(range, getFontHeightApplier(fontHeight));
	}
	
	/**
	 * Apply the font size to cells in the range, it will also enlarge the row height if row height is smaller than font size 
	 * @param range range to be applied
	 * @param point the font size in point
	 */
	public static void applyFontSize(Range range, final short point) {
		//fontSize = fontHeightInPoints = fontHeight/20
		applyFontHeight(range,(short)UnitUtil.pointToTwip(point));
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
		public void apply(Range cellRange,EditableCellStyle newCellstyle,EditableFont newfont);
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
				EditableCellStyle nstyle = cellRange.getCellStyleHelper().createCellStyle(ostyle);
				
				EditableFont nfont = null;
				//3.search if there any font already in book
				Font sfont = applyer.search(cellRange,ostyle,ofont);
				if( sfont== null){
					//create new font base old font if not found
					nfont = cellRange.getCellStyleHelper().createFont(ofont);
					nstyle.setFont(nfont);
					
					//set the apply
					applyer.apply(cellRange,nstyle, nfont);
				}else{
					nstyle.setFont(sfont);
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

	public static FontStyleApplier getFontBoldweightApplier(final Boldweight boldweight) {
		return  new FontStyleApplier() {
			public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
				return oldFont.getBoldweight().equals(boldweight);
			}
			
			public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
				return cellRange.getCellStyleHelper().findFont(boldweight, oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), 
						oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(Range range,EditableCellStyle newCellstyle, EditableFont newfont) {
				newfont.setBoldweight(boldweight);
			}
		};
	}
	
	/**
	 * Apply font bold-weight to cells in the range
	 * @param range the range to be applied
	 * @param boldweight the font bold-weight
	 */
	public static void applyFontBoldweight(Range range,final Boldweight boldweight) {
		applyFontStyle(range,getFontBoldweightApplier(boldweight));
	}

	public static FontStyleApplier getFontItalicApplier(final boolean italic) {
		return new FontStyleApplier() {
			public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
				return oldFont.isItalic()==italic;
			}
			
			public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
				return cellRange.getCellStyleHelper().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), 
						italic, oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(Range range,EditableCellStyle newCellstyle, EditableFont newfont) {
				newfont.setItalic(italic);
			}
		};
	}
	/**
	 * Apply font italic to cells in the range
	 * @param range the range to be applied
	 * @param italic the font italic
	 */
	public static void applyFontItalic(Range range, final boolean italic) {
		applyFontStyle(range, getFontItalicApplier(italic));
	}

	
	public static FontStyleApplier getFontStrikeoutApplier(final boolean strikeout) {
		return new FontStyleApplier() {
			public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
				return oldFont.isStrikeout()==strikeout;
			}
			
			public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
				return cellRange.getCellStyleHelper().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), 
						oldFont.isItalic(), strikeout, oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(Range range,EditableCellStyle newCellstyle, EditableFont newfont) {
				newfont.setStrikeout(strikeout);
			}
		};
	}
	/**
	 * Apply font strike-out to cells in the range 
	 * @param range the range to be applied
	 * @param strikeout font strike-out
	 */
	public static void applyFontStrikeout(Range range, final boolean strikeout) {
		applyFontStyle(range, getFontStrikeoutApplier(strikeout));
	}
	
	
	public static FontStyleApplier getFontUnderlineApplier(final Underline underline) {
		return new FontStyleApplier() {
			public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
				return oldFont.getUnderline().equals(underline);
			}
			
			public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
				return cellRange.getCellStyleHelper().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), 
						oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), underline);
			}
			
			public void apply(Range range,EditableCellStyle newCellstyle, EditableFont newfont) {
				newfont.setUnderline(underline);
			}
		};
	}
	
	/**
	 * Apply font underline to cells in the range
	 * @param range the range to be applied
	 * @param underline font underline
	 */
	public static void applyFontUnderline(Range range,final Underline underline) {
		applyFontStyle(range, getFontUnderlineApplier(underline));
	}
	
	public static FontStyleApplier getFontColorApplier(final Color color) {
		return new FontStyleApplier() {
			public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
				return oldFont.getColor().getHtmlColor().equals(color.getHtmlColor());
			}
			
			public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
				return cellRange.getCellStyleHelper().findFont(oldFont.getBoldweight(), color, oldFont.getFontHeight(), oldFont.getFontName(), 
						oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(Range range,EditableCellStyle newCellstyle, EditableFont newfont) {
				
				//check it in XSSFCellStyle , it is just a delegator, but do some thing in HSSF
				newfont.setColor(color);
				
				//TODO need to check with Henri, and Sam, why current implementation doesn't call set font color directly
				//TODO call style's set font color will cause set color after set a theme color issue(after clone form default)
//				newCellstyle.setFontColor(color); 
			}
		};
	}
	
	/**
	 * Apply font color to cells in the range
	 * @param range the range to be applied.
	 * @param htmlColor the color by html color syntax(#rgb-hex-code, e.x #FF00FF) 
	 */
	public static void applyFontColor(Range range, final String htmlColor) {
		final Color color = range.getCellStyleHelper().createColorFromHtmlColor(htmlColor);
		applyFontStyle(range, getFontColorApplier(color));
	}
	
	public static CellStyleApplier getBackgroundColorApplier(final Color color) {
		return new CellStyleApplier() {
				public boolean ignore(Range cellRange, CellStyle oldCellstyle) {
					Color ocolor = oldCellstyle.getBackgroundColor();
					return ocolor.equals(color);
				}

				public void apply(Range cellRange, EditableCellStyle newCellstyle) {
					newCellstyle.setBackgroundColor(color);
					
					FillPattern patternType = newCellstyle.getFillPattern();
					if (patternType == FillPattern.NO_FILL) {
						newCellstyle.setFillPattern(FillPattern.SOLID_FOREGROUND);
					}
				}
		};
	}
	
	/**
	 * Apply backgound-color to cells in the range
	 * @param range the range to be applied
	 * @param htmlColor the color by html color syntax(#rgb-hex-code, e.x #FF00FF) 
	 */
	public static void applyBackgroundColor(Range range, final String htmlColor) {
		final Color color = range.getCellStyleHelper().createColorFromHtmlColor(htmlColor);
		applyCellStyle(range,getBackgroundColorApplier(color));
	}
	
	
	public static CellStyleApplier getDataFormatApplier(final String format) {
		return new CellStyleApplier() {

			public boolean ignore(Range cellRange, CellStyle oldCellstyle) {
				String oformat = oldCellstyle.getDataFormat();
				return oformat.equals(format);
			}

			public void apply(Range cellRange, EditableCellStyle newCellstyle) {
				newCellstyle.setDataFormat(format);
			}
		};
	}
	
	/**
	 * Apply data-format to cells in the range
	 * @param range the range to be applied
	 * @param format the data format
	 */
	public static void applyDataFormat(Range range, final String format) {
		applyCellStyle(range, getDataFormatApplier(format));
	}

	public static CellStyleApplier getAligmentApplier(final Alignment alignment){
		return new CellStyleApplier() {

			public boolean ignore(Range cellRange, CellStyle oldCellstyle) {
				Alignment oldalign = oldCellstyle.getAlignment();
				return oldalign.equals(alignment);
			}

			public void apply(Range cellRange, EditableCellStyle newCellstyle) {
				newCellstyle.setAlignment(alignment);
			}
		};
	}
	
	/**
	 * Apply alignment to cells in the range
	 * @param range the range to be applied
	 * @param alignment the alignement
	 */
	public static void applyAlignment(Range range,final Alignment alignment) {
		applyCellStyle(range, getAligmentApplier(alignment));
	}

	public static CellStyleApplier getVerticalAligmentApplier(final VerticalAlignment alignment){
		return new CellStyleApplier() {

			public boolean ignore(Range cellRange, CellStyle oldCellstyle) {
				VerticalAlignment oldalign = oldCellstyle.getVerticalAlignment();
				return oldalign.equals(alignment);
			}

			public void apply(Range cellRange, EditableCellStyle newCellstyle) {
				newCellstyle.setVerticalAlignment(alignment);
			}
		};
	}
	
	/**
	 * Apply vertical-alignment to cells in the range
	 * @param range the range to be applied
	 * @param alignment vertical alignment
	 */
	public static void applyVerticalAlignment(Range range,final VerticalAlignment alignment) {
		applyCellStyle(range, getVerticalAligmentApplier(alignment));
	}
	
	/**
	 * Interface for help apply cell style
	 */
	public interface CellStyleApplier {
		/** should ignore this cellRange**/
		public boolean ignore(Range cellRange,CellStyle oldCellstyle);
		/** apply style to new cell**/
		public void apply(Range cellRange,EditableCellStyle newCellstyle);
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
				EditableCellStyle nstyle = cellRange.getCellStyleHelper().createCellStyle(
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
	 * @param applyType the apply type
	 * @param borderType the border type
	 * @param htmlColor the color of border(#rgb-hex-code, e.x #FF00FF) 
	 */
	public static void applyBorder(Range range,ApplyBorderType applyType,BorderType borderType,String htmlColor){
		if(range.isProtected())
			return;
		//use range api directly,
		range.applyBorders(applyType, borderType, htmlColor);
	}
	
	
	/**
	 * Toggle merge/unmerge of the range, if merging it will also set alignment to center
     * @param range the range to be applied
	 */
	public static void toggleMergeCenter(Range range){
		if(range.isProtected())
			return;
		range.sync(new RangeRunner() {
			public void run(Range range) {
				if(range.hasMergedCell()){
					range.unmerge();
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
	 * Unmerge the range
	 * @param range the range to be unmerge
	 */
	public static void unmerge(Range range){
		if(range.isProtected())
			return;
		range.unmerge();
	}
	
	public static CellStyleApplier getWrapTextApplier(final boolean wraptext) {
		return new CellStyleApplier() {

			public boolean ignore(Range cellRange, CellStyle oldCellstyle) {
				boolean oldwrap = oldCellstyle.isWrapText();
				return oldwrap==wraptext;
			}

			public void apply(Range cellRange, EditableCellStyle newCellstyle) {
				newCellstyle.setWrapText(wraptext);
			}
		};
	}
	
	/**
	 * Apply text-warp to cells in the range
     * @param range the range to be applied
	 * @param wraptext wrap text or not
	 */
	public static void applyWrapText(Range range,final boolean wraptext) {
		applyCellStyle(range, getWrapTextApplier(wraptext));
	}
	
//	public static CellStyleApplier getFormatApplier(final String format){
//		return new CellStyleApplier() {
//
//			public boolean ignore(Range cellRange, CellStyle oldCellstyle) {
//				String oformat = oldCellstyle.getDataFormat();
//				return Objects.equals(format, oformat);
//			}
//
//			public void apply(Range cellRange, EditableCellStyle newCellstyle) {
//				newCellstyle.setDataFormat(format);
//			}
//		};
//	}
//	
//	public static void applyFormat(Range range,final String format) {
//		applyCellStyle(range, getFormatApplier(format));
//	}
	
	
	public static void applyHyperlink(Range range,HyperlinkType type,String address,String label) {
		if(range.isProtected())
			return;
		
		range.setCellHyperlink(type, address, label);
		
		//make excel show link color style. 
		applyFontUnderline(range, Underline.SINGLE);
		applyFontColor(range, "#0000FF");
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
		range.unmerge(); // don't forge to unmerge the cell as well (ZSS-298)
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
	 * Insert rows to the range. 
	 * @param range the range to insert new rows 
	 */
	public static void insertRow(Range range) {
		insert(range.toRowRange(),InsertShift.DOWN,InsertCopyOrigin.FORMAT_LEFT_ABOVE);
	}
	
	/**
	 * Insert columns to the range. 
	 * @param range the range to insert new rows 
	 */
	public static void insertColumn(Range range) {
		insert(range.toColumnRange(),InsertShift.RIGHT,InsertCopyOrigin.FORMAT_LEFT_ABOVE);
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
	 * Delete rows of the range. 
	 * @param range the range to delete rows 
	 */
	public static void deleteRow(Range range) {
		delete(range.toRowRange(),DeleteShift.UP);
	}
	
	/**
	 * Delete columns to the range. 
	 * @param range the range to delete columns 
	 */
	public static void deleteColumn(Range range) {
		delete(range.toColumnRange(),DeleteShift.LEFT);
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
	 * Sort range
	 * @param range the range to sort
	 * @param desc true for descent, false for ascent
	 */
	public static void sort(Range range,Range index1,boolean desc1,SortDataOption dataOption1,
			Range index2,boolean desc2,SortDataOption dataOption2,
			Range index3,boolean desc3,SortDataOption dataOption3,
			boolean header,
			boolean matchCase, 
			boolean sortByRows 
			) {
		if(range.isProtected())
			return;
		range.sort(index1,desc1,dataOption1,index2,desc2,dataOption2,index3,desc3,dataOption3,header,matchCase,sortByRows);
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
	public static void unhide(Range range) {
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
	
	/**
	 * Fills data from source range to destination range automatically upon auto fill type
	 * @param src the source range
	 * @param dest the destination range
	 * @param type the fill type, currently only support AutoFillType.DEFAULT, AutoFillType.COPY, AutoFillType.FORMAT, AutoFillType.VALUES
	 */
	public static void autoFill(Range src, Range dest, AutoFillType type) {
		if(dest.isProtected())
			return;
		src.autoFill(dest, type);
	}
	
	public static void setRowHeight(Range range, int heightPx) {
		range.setRowHeight(heightPx);
	}
	
	public static void setColumnWidth(Range range, int widthPx) {
		range.setColumnWidth(widthPx);
	}
}
