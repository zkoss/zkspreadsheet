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

import java.util.HashMap;

import org.zkoss.util.Locales;
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
import org.zkoss.zss.api.model.impl.EnumUtil;
import org.zkoss.zss.api.model.impl.SheetImpl;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.sys.EngineFactory;
import org.zkoss.zss.model.sys.format.FormatContext;
import org.zkoss.zss.range.impl.StyleUtil;

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

	public static CellStyleApplier getFontNameApplier(final String fontName){
		return new CellStyleApplier() {
			public void apply(Range range) {
				//ZSS 464, efficient implementation
				StyleUtil.setFontName(((SheetImpl)range.getSheet()).getNative(),range.getRow(),range.getColumn(),fontName);
			}
		};
	}
	/**
	 * Apply font to cells in the range
	 * @param range range to be applied
	 * @param fontName the font name
	 */
	public static void applyFontName(Range range,final String fontName){
		applyCellStyle(range, getFontNameApplier(fontName));
	}
	
	/**
	 * @deprecated use {@link #applyFontHeightPoints(Range, int)}
	 */
	public static CellStyleApplier getFontHeightApplier(final int fontHeight) {
		//fontHeightPoints = pt;
		return getFontHeightPointsApplier(UnitUtil.twipToPoint(fontHeight));
	}

	/**
	 * Apply font height to cells in the range, it will also enlarge the row height if row height is smaller than font height 
	 * @param range range to be applied
	 * @param fontHeight the font height in twpi (1/20 point)
	 * @deprecated use {@link #applyFontHeightPoints(Range, int)}
	 * 
	 */
	public static void applyFontHeight(Range range, final int fontHeight) {
		applyCellStyle(range, getFontHeightApplier(fontHeight));
	}
	
	/**
	 * Apply the font size to cells in the range, it will also enlarge the row height if row height is smaller than font size 
	 * @param range range to be applied
	 * @param point the font size in point
	 * @deprecated use {@link #applyFontHeightPoints(Range, int)}
	 */
	public static void applyFontSize(Range range, final int point) {
		//fontSize = fontHeightInPoints = fontHeight/20
		applyFontHeightPoints(range,point);
	}

	/**
	 * Apply font height to cells in the range, it will also enlarge the row height if row height is smaller than font height 
	 * @param range range to be applied
	 * @param fontHeightPoints the font height in point
	 */
	public static void applyFontHeightPoints(Range range, final int fontHeightPoints) {
		applyCellStyle(range, getFontHeightPointsApplier(fontHeightPoints));
	}

	public static CellStyleApplier getFontHeightPointsApplier(final int fontHeightPoints) {
		//fontHeightPoints = pt;
		final int fpx = UnitUtil.pointToPx(fontHeightPoints);
		return new CellStyleApplier() {
			public void apply(Range range) {
				StyleUtil.setFontHeightPoints(((SheetImpl)range.getSheet()).getNative(),range.getRow(),range.getColumn(),fontHeightPoints);
				int px = range.getSheet().getRowHeight(range.getRow());//rowHeight in px
				if(fpx>px){
					range.setRowHeight(fpx+4);//4 is padding
				}
			}
		};
	}

	public static CellStyleApplier getFontBoldweightApplier(final Boldweight boldweight) {
		return  new CellStyleApplier() {
			public void apply(Range range) {
				//ZSS 464, efficient implement
				StyleUtil.setFontBoldWeight(((SheetImpl)range.getSheet()).getNative(),range.getRow(),range.getColumn(),EnumUtil.toFontBoldweight(boldweight));
			}
		};
	}
	
	/**
	 * Apply font bold-weight to cells in the range
	 * @param range the range to be applied
	 * @param boldweight the font bold-weight
	 */
	public static void applyFontBoldweight(Range range,final Boldweight boldweight) {
		applyCellStyle(range,getFontBoldweightApplier(boldweight));
	}

	public static CellStyleApplier getFontItalicApplier(final boolean italic) {
		return new CellStyleApplier() {
			public void apply(Range range) {
				//ZSS 464, efficient implement
				StyleUtil.setFontItalic(((SheetImpl)range.getSheet()).getNative(),range.getRow(),range.getColumn(),italic);
			}
		};
	}
	/**
	 * Apply font italic to cells in the range
	 * @param range the range to be applied
	 * @param italic the font italic
	 */
	public static void applyFontItalic(Range range, final boolean italic) {
		applyCellStyle(range, getFontItalicApplier(italic));
	}

	
	public static CellStyleApplier getFontStrikeoutApplier(final boolean strikeout) {
		return new CellStyleApplier() {
			public void apply(Range range) {
				//ZSS 464, efficient implement
				StyleUtil.setFontStrikethrough(((SheetImpl)range.getSheet()).getNative(),range.getRow(),range.getColumn(),strikeout);
			}
		};
	}
	/**
	 * Apply font strike-out to cells in the range 
	 * @param range the range to be applied
	 * @param strikeout font strike-out
	 */
	public static void applyFontStrikeout(Range range, final boolean strikeout) {
		applyCellStyle(range, getFontStrikeoutApplier(strikeout));
	}
	
	
	public static CellStyleApplier getFontUnderlineApplier(final Underline underline) {
		return new CellStyleApplier() {
			public void apply(Range range) {
				//ZSS 464, efficient implement
				StyleUtil.setFontUnderline(((SheetImpl)range.getSheet()).getNative(),range.getRow(),range.getColumn(),EnumUtil.toFontUnderline(underline));
			}
		};
	}
	
	/**
	 * Apply font underline to cells in the range
	 * @param range the range to be applied
	 * @param underline font underline
	 */
	public static void applyFontUnderline(Range range,final Underline underline) {
		applyCellStyle(range, getFontUnderlineApplier(underline));
	}
	
	public static CellStyleApplier getFontColorApplier(final Color color) {
		return new CellStyleApplier() {
//			private HashMap<Integer,org.zkoss.poi.ss.usermodel.CellStyle> _cache = new HashMap<Integer,org.zkoss.poi.ss.usermodel.CellStyle>();
			public void apply(Range range) {
				//ZSS 464, efficient implement
				StyleUtil.setFontColor(((SheetImpl)range.getSheet()).getNative(),range.getRow(),range.getColumn(),color.getHtmlColor());
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
		applyCellStyle(range, getFontColorApplier(color));
	}
	
	@Deprecated
	public static CellStyleApplier getBackgroundColorApplier(final Color color) {
		return getFillColorApplier(color);
	}
	
	public static CellStyleApplier getFillColorApplier(final Color color) {
		return new CellStyleApplier() {
				public void apply(Range range) {
					StyleUtil.setFillColor(((SheetImpl)range.getSheet()).getNative(),range.getRow(),range.getColumn(),color.getHtmlColor());
				}
		};
	}
	
	/**
	 * Apply backgound-color to cells in the range
	 * @param range the range to be applied
	 * @param htmlColor the color by html color syntax(#rgb-hex-code, e.x #FF00FF)
	 * @Deprecated since 3.5.0 use {@link #applyFillColor(Range, String)}
	 */
	@Deprecated
	public static void applyBackgroundColor(Range range, final String htmlColor) {
		applyFillColor(range,htmlColor);
	}
	
	/**
	 * Apply fill-color to cells in the range
	 * @param range the range to be applied
	 * @param htmlColor the color by html color syntax(#rgb-hex-code, e.x #FF00FF) 
	 */
	public static void applyFillColor(Range range, final String htmlColor) {
		final Color color = range.getCellStyleHelper().createColorFromHtmlColor(htmlColor);
		applyCellStyle(range,getFillColorApplier(color));
	}	
	
	
	public static CellStyleApplier getDataFormatApplier(final String format) {
		return new CellStyleApplier() {
			public void apply(Range range) {
				//ZSS 464, efficient implement
				StyleUtil.setDataFormat(((SheetImpl)range.getSheet()).getNative(),range.getRow(),range.getColumn(),format);
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
			public void apply(Range range) {
				//ZSS 464, efficient implement
				StyleUtil.setTextHAlign(((SheetImpl)range.getSheet()).getNative(),range.getRow(),range.getColumn(),EnumUtil.toStyleAlignemnt(alignment));
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
			public void apply(Range range) {
				//ZSS 464, efficient implement
				StyleUtil.setTextVAlign(((SheetImpl)range.getSheet()).getNative(),range.getRow(),range.getColumn(),EnumUtil.toStyleVerticalAlignemnt(alignment));
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
		/** apply style to new cell**/
		public void apply(Range cellRange);
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
				// set the apply
				applyer.apply(cellRange);
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
		range.notifyChange();
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
			public void apply(Range range) {
				StyleUtil.setTextWrap(((SheetImpl)range.getSheet()).getNative(),range.getRow(),range.getColumn(),wraptext);
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
		range.clearAll();
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
	
	/**
	 * Sets the row height and make it as custom modified
	 * @param range
	 * @param heightPx
	 */
	public static void setRowHeight(Range range, int heightPx) {
		range.setRowHeight(heightPx);
	}
	
	/**
	 * Sets the row height and provide a custom flag, a custom flag to indicate this height was set by user or system. 
	 * @param range
	 * @param heightPx
	 * @param isCustom
	 * @since 3.0.1
	 */
	public static void setRowHeight(Range range, int heightPx, boolean isCustom) {
		range.setRowHeight(heightPx, isCustom);
	}

	/**
	 * Sets the column width
	 * @param range
	 * @param widthPx
	 */
	public static void setColumnWidth(Range range, int widthPx) {
		range.setColumnWidth(widthPx);
	}
}
