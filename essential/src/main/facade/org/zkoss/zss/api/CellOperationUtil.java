package org.zkoss.zss.api;

import org.zkoss.zss.api.NRange.BatchLockLevel;
import org.zkoss.zss.api.NRange.PasteOperation;
import org.zkoss.zss.api.NRange.PasteType;
import org.zkoss.zss.api.NRange.VisitorLockLevel;
import org.zkoss.zss.api.model.NCellStyle;
import org.zkoss.zss.api.model.NCellStyle.Alignment;
import org.zkoss.zss.api.model.NCellStyle.FillPattern;
import org.zkoss.zss.api.model.NCellStyle.VerticalAlignment;
import org.zkoss.zss.api.model.NColor;
import org.zkoss.zss.api.model.NFont;
import org.zkoss.zss.api.model.NFont.Boldweight;
import org.zkoss.zss.api.model.NFont.Underline;

public class CellOperationUtil {

	public static boolean cut(NRange src, final NRange dest) {
		final Result<Boolean> result = new Result<Boolean>();
		//use batch-runner to run multiple range operation
		src.batch(new NBatchRunner() {
			public void run(NRange range) {
				boolean r = range.paste(dest);
				if(r){
					range.clearContents();// it removes value and formula only
					range.clearStyles();
				}
				result.set(r);
			}
		}, BatchLockLevel.SHEET);

		return result.get();
	}

	public static boolean paste(NRange src, NRange dest) {
		return src.paste(dest);
	}
	public static boolean pasteFormula(NRange src, NRange dest) {
		return src.pasteSpecial(dest, PasteType.PASTE_FORMULAS, PasteOperation.PASTEOP_NONE, false, false);
	}
	public static boolean pasteValue(NRange src, NRange dest) {
		return src.pasteSpecial(dest, PasteType.PASTE_VALUES, PasteOperation.PASTEOP_NONE, false, false);
	}
	public static boolean pasteAllExceptBorder(NRange src, NRange dest) {
		return src.pasteSpecial(dest, PasteType.PASTE_ALL_EXCEPT_BORDERS, PasteOperation.PASTEOP_NONE, false, false);
	}
	public static boolean pasteTranspose(NRange src, NRange dest) {
		return src.pasteSpecial(dest, PasteType.PASTE_ALL, PasteOperation.PASTEOP_NONE, false, true);
	}
	
	static class Result<T> {
		T r;
		public Result(){}
		public Result(T r){
			this.r = r;
		}
		
		public T get(){
			return r;
		}
		
		public void set(T r){
			this.r = r;
		}
	}
	
	public static void applyFontName(NRange range,final String fontName){
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(NRange range,NCellStyle oldCellstyle, NFont oldFont) {
				//the font name(family) is equals, not need to set it.
				return oldFont.getFontName().equals(fontName);
			}
			
			public NFont search(NRange cellRange,NCellStyle oldCellstyle, NFont oldFont){
				//use the new font name to search it.
				return cellRange.getBook().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), fontName, 
						oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(NRange range,NCellStyle cellstyle, NFont newfont) {
				newfont.setFontName(fontName);
			}
		});
	}

	public static void applyFontHeight(NRange range, final short height) {
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(NRange range,NCellStyle oldCellstyle, NFont oldFont) {
				//the font name(family) is equals, not need to set it.
				return oldFont.getFontHeight() == height;
			}
			
			public NFont search(NRange cellRange,NCellStyle oldCellstyle, NFont oldFont){
				//use the new font name to search it.
				return cellRange.getBook().findFont(oldFont.getBoldweight(), oldFont.getColor(), height, oldFont.getFontName(), 
						oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(NRange range,NCellStyle cellstyle, NFont newfont) {
				newfont.setFontHeight(height);
			}
		});
	}
	public static void applyFontSize(NRange range, final short size) {
		//fontSize = fontHeightInPoints = fontHeight/20
		applyFontHeight(range,(short)(size*20));
	}
	
	public interface FontStyleApplier {
		/** should ignore this cellRange**/
		public boolean ignore(NRange cellRange,NCellStyle oldCellstyle,NFont oldFont);
		/** find the font to apply to new style, return null mean not found, and will create a new font**/
		public NFont search(NRange cellRange,NCellStyle oldCellstyle, NFont oldFont);
		/** apply style to new font, will be call when {@code #search() return null}**/
		public void apply(NRange cellRange,NCellStyle newCellstyle,NFont newfont);
	}
	
	public static void applyFontStyle(NRange range,final FontStyleApplier applyer){
		//use use cell visitor to visit cells respectively
		//use BOOK level lock because we will create BOOK level Object (style, font)
		range.visit(new NCellVisitor(){
			@Override
			public void visit(NRange cellRange) {
				if(cellRange.isAnyCellProtected()){//don't apply if protected
					return;
				}
				NCellStyle ostyle = cellRange.getGetter().getCellStyle();
				NFont ofont = ostyle.getFont();
				if(applyer.ignore(cellRange,ostyle,ofont)){//ignore or not, to prevent unnecessary creation
					return;
				}
				//2.create a new style from old one, the original one is shared between cells
				//TODO what if it is the last referenced, could I just use it? who will delete old if I use new
				NCellStyle nstyle = cellRange.getCreator().createCellStyle(ostyle);
				
				NFont nfont = null;
				//3.search if there any font already in book
				nfont = applyer.search(cellRange,ostyle,ofont);
				if( nfont== null){
					//create new font base old font if not found
					nfont = cellRange.getCreator().createFont(ofont);
					nstyle.setFont(nfont);
					
					//set the apply
					applyer.apply(cellRange,nstyle, nfont);
				}else{
					nstyle.setFont(nfont);
				}
				
				
				cellRange.setStyle(nstyle);
			}},VisitorLockLevel.BOOK);
	}

	

	public static void applyFontBoldweight(NRange range,final Boldweight boldweight) {
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(NRange range,NCellStyle oldCellstyle, NFont oldFont) {
				return oldFont.getBoldweight().equals(boldweight);
			}
			
			public NFont search(NRange cellRange,NCellStyle oldCellstyle, NFont oldFont){
				return cellRange.getBook().findFont(boldweight, oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), 
						oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(NRange range,NCellStyle newCellstyle, NFont newfont) {
				newfont.setBoldweight(boldweight);
			}
		});
	}

	public static void applyFontItalic(NRange range, final boolean italic) {
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(NRange range,NCellStyle oldCellstyle, NFont oldFont) {
				return oldFont.isItalic()==italic;
			}
			
			public NFont search(NRange cellRange,NCellStyle oldCellstyle, NFont oldFont){
				return cellRange.getBook().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), 
						italic, oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(NRange range,NCellStyle newCellstyle, NFont newfont) {
				newfont.setItalic(italic);
			}
		});
	}

	public static void applyFontStrikeout(NRange range, final boolean strikeout) {
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(NRange range,NCellStyle oldCellstyle, NFont oldFont) {
				return oldFont.isStrikeout()==strikeout;
			}
			
			public NFont search(NRange cellRange,NCellStyle oldCellstyle, NFont oldFont){
				return cellRange.getBook().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), 
						oldFont.isItalic(), strikeout, oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(NRange range,NCellStyle newCellstyle, NFont newfont) {
				newfont.setStrikeout(strikeout);
			}
		});
	}
	
	public static void applyFontUnderline(NRange range,final Underline underline) {
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(NRange range,NCellStyle oldCellstyle, NFont oldFont) {
				return oldFont.getUnderline().equals(underline);
			}
			
			public NFont search(NRange cellRange,NCellStyle oldCellstyle, NFont oldFont){
				return cellRange.getBook().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), 
						oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), underline);
			}
			
			public void apply(NRange range,NCellStyle newCellstyle, NFont newfont) {
				newfont.setUnderline(underline);
			}
		});
	}
	
	/**
	 * @param htmlColor '#rgb-hex-code'
	 */
	public static void applyFontColor(NRange range, final String htmlColor) {
		final NColor color = range.getBook().getColorFromHtmlColor(htmlColor);
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(NRange range,NCellStyle oldCellstyle, NFont oldFont) {
				return oldFont.getColor().toHtmlColor().equals(htmlColor);
			}
			
			public NFont search(NRange cellRange,NCellStyle oldCellstyle, NFont oldFont){
				return cellRange.getBook().findFont(oldFont.getBoldweight(), color, oldFont.getFontHeight(), oldFont.getFontName(), 
						oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(NRange range,NCellStyle newCellstyle, NFont newfont) {
				
				//check it in XSSFCellStyle , it is just a delegator, but do some thing in HSSF
				newfont.setFontColor(color);
				
				//TODO need to check with Henri, and Sam, why current implementation doesn't call set font color directly
				//TODO call style's set font color will cause set color after set a theme color issue(after clone form default)
//				newCellstyle.setFontColor(color); 
			}
		});
	}
	
	
	/**
	 * @param htmlColor '#rgb-hex-code'
	 */
	public static void applyCellColor(NRange range, final String htmlColor) {
		final NColor color = range.getBook().getColorFromHtmlColor(htmlColor);
		applyCellStyle(range, new CellStyleApplier() {

			public boolean ignore(NRange cellRange, NCellStyle oldCellstyle) {
				NColor ocolor = oldCellstyle.getBackgroundColor();
				return ocolor.equals(color);
			}

			public void apply(NRange cellRange, NCellStyle newCellstyle) {
				newCellstyle.setBackgroundColor(color);
				
				FillPattern patternType = newCellstyle.getFillPattern();
				if (patternType == FillPattern.NO_FILL) {
					newCellstyle.setFillPattern(FillPattern.SOLID_FOREGROUND);
				}
			}
		});
	}

	public static void applyCellAlignment(NRange range,final Alignment alignment) {
		applyCellStyle(range, new CellStyleApplier() {

			public boolean ignore(NRange cellRange, NCellStyle oldCellstyle) {
				Alignment oldalign = oldCellstyle.getAlignment();
				return oldalign.equals(alignment);
			}

			public void apply(NRange cellRange, NCellStyle newCellstyle) {
				newCellstyle.setAlignment(alignment);
			}
		});
	}

	public static void applyCellVerticalAlignment(NRange range,final VerticalAlignment alignment) {
		applyCellStyle(range, new CellStyleApplier() {

			public boolean ignore(NRange cellRange, NCellStyle oldCellstyle) {
				VerticalAlignment oldalign = oldCellstyle.getVerticalAlignment();
				return oldalign.equals(alignment);
			}

			public void apply(NRange cellRange, NCellStyle newCellstyle) {
				newCellstyle.setVerticalAlignment(alignment);
			}
		});
	}
	
	
	public interface CellStyleApplier {
		/** should ignore this cellRange**/
		public boolean ignore(NRange cellRange,NCellStyle oldCellstyle);
		/** apply style to new cell**/
		public void apply(NRange cellRange,NCellStyle newCellstyle);
	}
	
	public static void applyCellStyle(NRange range,
			final CellStyleApplier applyer) {
		// use use cell visitor to visit cells respectively
		// use BOOK level lock because we will create BOOK level Object (style,
		// font)
		range.visit(new NCellVisitor() {
			@Override
			public void visit(NRange cellRange) {
				if (cellRange.isAnyCellProtected()) {// don't apply if protected
					return;
				}
				NCellStyle ostyle = cellRange.getGetter().getCellStyle();
				if (applyer.ignore(cellRange, ostyle)) {// ignore or not, to
														// prevent unnecessary
														// creation
					return;
				}
				// 2.create a new style from old one, the original one is shared
				// between cells
				// TODO what if it is the last referenced, could I just use it?
				// who will delete old if I use new
				NCellStyle nstyle = cellRange.getCreator().createCellStyle(
						ostyle);

				// set the apply
				applyer.apply(cellRange, nstyle);

				cellRange.setStyle(nstyle);
			}
		}, VisitorLockLevel.BOOK);
	}
}
