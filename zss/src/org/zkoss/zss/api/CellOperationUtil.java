package org.zkoss.zss.api;

import org.zkoss.image.AImage;
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
import org.zkoss.zss.api.model.Chart;
import org.zkoss.zss.api.model.ChartData;
import org.zkoss.zss.api.model.Color;
import org.zkoss.zss.api.model.Font;
import org.zkoss.zss.api.model.Font.Boldweight;
import org.zkoss.zss.api.model.Font.Underline;
import org.zkoss.zss.api.model.Picture.Format;

/**
 * the utit to help UI to deal with UI operation of a Range.
 * it also handles 1.the sheet protection, 2. 
 * @author dennis
 *
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

	public static boolean paste(Range src, Range dest) {
		if(dest.isProtected()){
			return false;
		}
		return src.paste(dest);
	}
	public static boolean pasteFormula(Range src, Range dest) {
		return pasteSpecial(src, dest, PasteType.PASTE_FORMULAS, PasteOperation.PASTEOP_NONE, false, false);
	}
	public static boolean pasteValue(Range src, Range dest) {
		return pasteSpecial(src, dest, PasteType.PASTE_VALUES, PasteOperation.PASTEOP_NONE, false, false);
	}
	public static boolean pasteAllExceptBorder(Range src, Range dest) {
		return pasteSpecial(src, dest, PasteType.PASTE_ALL_EXCEPT_BORDERS, PasteOperation.PASTEOP_NONE, false, false);
	}
	public static boolean pasteTranspose(Range src, Range dest) {
		return pasteSpecial(src, dest, PasteType.PASTE_ALL, PasteOperation.PASTEOP_NONE, false, true);
	}
	
	public static boolean pasteSpecial(Range src, Range dest,PasteType pasteType, PasteOperation pasteOperation, boolean skipBlank, boolean transpose){
		if(dest.isProtected()){
			return false;
		}
		return src.pasteSpecial(dest, pasteType, pasteOperation, skipBlank, transpose);
	}

	public static void applyFontName(Range range,final String fontName){
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
				//the font name(family) is equals, not need to set it.
				return oldFont.getFontName().equals(fontName);
			}
			
			public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
				//use the new font name to search it.
				return cellRange.getStyleHelper().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), fontName, 
						oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(Range range,CellStyle cellstyle, Font newfont) {
				newfont.setFontName(fontName);
			}
		});
	}

	public static void applyFontHeight(Range range, final short height) {
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
				//the font name(family) is equals, not need to set it.
				return oldFont.getFontHeight() == height;
			}
			
			public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
				//use the new font name to search it.
				return cellRange.getStyleHelper().findFont(oldFont.getBoldweight(), oldFont.getColor(), height, oldFont.getFontName(), 
						oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(Range range,CellStyle cellstyle, Font newfont) {
				newfont.setFontHeight(height);
			}
		});
	}
	public static void applyFontSize(Range range, final short size) {
		//fontSize = fontHeightInPoints = fontHeight/20
		applyFontHeight(range,(short)(size*20));
	}
	
	public interface FontStyleApplier {
		/** should ignore this cellRange**/
		public boolean ignore(Range cellRange,CellStyle oldCellstyle,Font oldFont);
		/** find the font to apply to new style, return null mean not found, and will create a new font**/
		public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont);
		/** apply style to new font, will be call when {@code #search() return null}**/
		public void apply(Range cellRange,CellStyle newCellstyle,Font newfont);
	}
	
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
				CellStyle nstyle = cellRange.getStyleHelper().createCellStyle(ostyle);
				
				Font nfont = null;
				//3.search if there any font already in book
				nfont = applyer.search(cellRange,ostyle,ofont);
				if( nfont== null){
					//create new font base old font if not found
					nfont = cellRange.getStyleHelper().createFont(ofont);
					nstyle.setFont(nfont);
					
					//set the apply
					applyer.apply(cellRange,nstyle, nfont);
				}else{
					nstyle.setFont(nfont);
				}
				
				
				cellRange.setStyle(nstyle);
				return true;
			}

			@Override
			public boolean ignoreIfNotExist(int row, int column) {
				return false;
			}});
	}

	

	public static void applyFontBoldweight(Range range,final Boldweight boldweight) {
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
				return oldFont.getBoldweight().equals(boldweight);
			}
			
			public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
				return cellRange.getStyleHelper().findFont(boldweight, oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), 
						oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(Range range,CellStyle newCellstyle, Font newfont) {
				newfont.setBoldweight(boldweight);
			}
		});
	}

	public static void applyFontItalic(Range range, final boolean italic) {
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
				return oldFont.isItalic()==italic;
			}
			
			public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
				return cellRange.getStyleHelper().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), 
						italic, oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(Range range,CellStyle newCellstyle, Font newfont) {
				newfont.setItalic(italic);
			}
		});
	}

	public static void applyFontStrikeout(Range range, final boolean strikeout) {
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
				return oldFont.isStrikeout()==strikeout;
			}
			
			public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
				return cellRange.getStyleHelper().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), 
						oldFont.isItalic(), strikeout, oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(Range range,CellStyle newCellstyle, Font newfont) {
				newfont.setStrikeout(strikeout);
			}
		});
	}
	
	public static void applyFontUnderline(Range range,final Underline underline) {
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
				return oldFont.getUnderline().equals(underline);
			}
			
			public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
				return cellRange.getStyleHelper().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), 
						oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), underline);
			}
			
			public void apply(Range range,CellStyle newCellstyle, Font newfont) {
				newfont.setUnderline(underline);
			}
		});
	}
	
	/**
	 * @param htmlColor '#rgb-hex-code'
	 */
	public static void applyFontColor(Range range, final String htmlColor) {
		final Color color = range.getStyleHelper().createColorFromHtmlColor(htmlColor);
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(Range range,CellStyle oldCellstyle, Font oldFont) {
				return oldFont.getColor().getHtmlColor().equals(htmlColor);
			}
			
			public Font search(Range cellRange,CellStyle oldCellstyle, Font oldFont){
				return cellRange.getStyleHelper().findFont(oldFont.getBoldweight(), color, oldFont.getFontHeight(), oldFont.getFontName(), 
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
	 * @param htmlColor '#rgb-hex-code'
	 */
	public static void applyCellColor(Range range, final String htmlColor) {
		final Color color = range.getStyleHelper().createColorFromHtmlColor(htmlColor);
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

	public static void applyCellAlignment(Range range,final Alignment alignment) {
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

	public static void applyCellVerticalAlignment(Range range,final VerticalAlignment alignment) {
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
	
	
	public interface CellStyleApplier {
		/** should ignore this cellRange**/
		public boolean ignore(Range cellRange,CellStyle oldCellstyle);
		/** apply style to new cell**/
		public void apply(Range cellRange,CellStyle newCellstyle);
	}
	
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
				CellStyle nstyle = cellRange.getStyleHelper().createCellStyle(
						ostyle);

				// set the apply
				applyer.apply(cellRange, nstyle);

				cellRange.setStyle(nstyle);
				return true;
			}

			@Override
			public boolean ignoreIfNotExist(int row, int column) {
				return false;
			}
		});
	}
	
	public static void applyBorder(Range range,ApplyBorderType type,BorderType borderType,String htmlColor){
		if(range.isProtected())
			return;
		//use range api directly,
		range.applyBorders(type, borderType, htmlColor);
	}
	
	
	public static void toggleMergeCenter(Range range){
		if(range.isProtected())
			return;
		range.sync(new RangeRunner() {
			public void run(Range range) {
				if(range.hasMergeCell()){
					range.unMerge();
				}else{
					range.merge(false);
					//align the left/top one
					applyCellAlignment(range.getCellRange(0,0),Alignment.CENTER);
				}
			}
		});
	}
	
	public static void merge(Range range,boolean across){
		if(range.isProtected())
			return;
		
		range.merge(across);
	}
	
	public static void unMerge(Range range){
		if(range.isProtected())
			return;
		range.unMerge();
	}
	
	
	public static void applyCellWrapText(Range range,final boolean wraptext) {
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

	public static void clearContents(Range range) {
		if(range.isProtected())
			return;
		range.clearContents();
	}
	
	public static void clearStyles(Range range) {
		if(range.isProtected())
			return;
		range.clearStyles();
	}
	
	public static void clearAll(Range range){
		if(range.isProtected())
			return;

		//use batch-runner to run multiple range operation
		range.sync(new RangeRunner() {
			public void run(Range range) {
				range.clearContents();// it removes value and formula only
				range.clearStyles();
			}
		});
	}

	public static void insert(Range range, InsertShift shift,
			InsertCopyOrigin copyOrigin) {
		if(range.isProtected())
			return;
		range.insert(shift, copyOrigin);
	}
	
	public static void delete(Range range, DeleteShift shift) {
		if(range.isProtected())
			return;
		range.delete(shift);
	}

	public static void sort(Range range, boolean desc) {
		if(range.isProtected())
			return;
		range.sort(desc);
	}

	public static void hide(Range range) {
		if (range.isProtected())
			return;
		range.setHidden(true);
	}
	public static void unHide(Range range) {
		if (range.isProtected())
			return;
		range.setHidden(false);
	}
}
