package org.zkoss.zss.api;

import org.zkoss.zss.api.NRange.BatchLockLevel;
import org.zkoss.zss.api.NRange.PasteOperation;
import org.zkoss.zss.api.NRange.PasteType;
import org.zkoss.zss.api.NRange.VisitorLockLevel;
import org.zkoss.zss.api.model.NCellStyle;
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
				return cellRange.getGetter().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), fontName, 
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
				return cellRange.getGetter().findFont(oldFont.getBoldweight(), oldFont.getColor(), height, oldFont.getFontName(), 
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
	
	private interface FontStyleApplier {
		public boolean ignore(NRange cellRange,NCellStyle oldCellstyle,NFont oldFont);
		public NFont search(NRange cellRange,NCellStyle oldCellstyle, NFont oldFont);
		public void apply(NRange cellRange,NCellStyle cellstyle,NFont newfont);
	}
	
	private static void applyFontStyle(NRange range,final FontStyleApplier applyer){
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
					//set the apply
					applyer.apply(cellRange,nstyle, nfont);
				}
				
				nstyle.setFont(nfont);
				cellRange.setStyle(nstyle);
			}},VisitorLockLevel.BOOK);
	}

	

	public static void applyFontBoldweight(NRange range,final Boldweight boldweight) {
		applyFontStyle(range, new FontStyleApplier() {
			public boolean ignore(NRange range,NCellStyle oldCellstyle, NFont oldFont) {
				return oldFont.getBoldweight().equals(boldweight);
			}
			
			public NFont search(NRange cellRange,NCellStyle oldCellstyle, NFont oldFont){
				return cellRange.getGetter().findFont(boldweight, oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), 
						oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(NRange range,NCellStyle cellstyle, NFont newfont) {
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
				return cellRange.getGetter().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), 
						italic, oldFont.isStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(NRange range,NCellStyle cellstyle, NFont newfont) {
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
				return cellRange.getGetter().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), 
						oldFont.isItalic(), strikeout, oldFont.getTypeOffset(), oldFont.getUnderline());
			}
			
			public void apply(NRange range,NCellStyle cellstyle, NFont newfont) {
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
				return cellRange.getGetter().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), 
						oldFont.isItalic(), oldFont.isStrikeout(), oldFont.getTypeOffset(), underline);
			}
			
			public void apply(NRange range,NCellStyle cellstyle, NFont newfont) {
				newfont.setUnderline(underline);
			}
		});
	}

}
