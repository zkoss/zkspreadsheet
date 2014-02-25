/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
 */
package org.zkoss.zss.model;

/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public interface SPrintSetup {
	
	public enum PaperSize {
	    /** US Letter 8 1/2 x 11 in */
	    LETTER,
	    /** US Letter Small 8 1/2 x 11 in */
	    LETTER_SMALL,
	    /** US Tabloid 11 x 17 in */
	    TABLOID,
	    /** US Ledger 17 x 11 in */
	    LEDGER,
	    /** US Legal 8 1/2 x 14 in */
	    LEGAL,
	    /** US Statement 5 1/2 x 8 1/2 in */
	    STATEMENT,
	    /** US Executive 7 1/4 x 10 1/2 in */
	    EXECUTIVE,
	    /** A3 - 297x420 mm */
	    A3,
	    /** A4 - 210x297 mm */
	    A4,
	    /** A4 Small - 210x297 mm */
	    A4_SMALL,
	    /** A5 - 148x210 mm */
	    A5,
	    /** B4 (JIS) 250x354 mm */
	    B4,
	    /** B5 (JIS) 182x257 mm */
	    B5,
	    /** Folio 8 1/2 x 13 in */
	    FOLIO8,
	    /** Quarto 215x275 mm */
	    QUARTO,
	    /** 10 x 14 in */
	    TEN_BY_FOURTEEN,
	    /** 11 x 17 in */
	    ELEVEN_BY_SEVENTEEN,
	    /** US Note 8 1/2 x 11 in */
	    NOTE8,
	    /** US Envelope #9 3 7/8 x 8 7/8 */
	    ENVELOPE_9,
	    /** US Envelope #10 4 1/8 x 9 1/2 */
	    ENVELOPE_10,
	    /** Envelope DL 110x220 mm */
	    ENVELOPE_DL,
	    /** Envelope C5 162x229 mm */
	    ENVELOPE_CS,
	    ENVELOPE_C5,
	    /** Envelope C3 324x458 mm */
	    ENVELOPE_C3,
	    /** Envelope C4 229x324 mm */
	    ENVELOPE_C4,
	    /** Envelope C6 114x162 mm */
	    ENVELOPE_C6,

	    ENVELOPE_MONARCH,
	    /** A4 Extra - 9.27 x 12.69 in */
	    A4_EXTRA,
	    /** A4 Transverse - 210x297 mm */
	    A4_TRANSVERSE,
	    /** A4 Plus - 210x330 mm */
	    A4_PLUS,
	    /** US Letter Rotated 11 x 8 1/2 in */
	    LETTER_ROTATED,
	    /** A4 Rotated - 297x210 mm */
	    A4_ROTATED
	}
	
	public boolean isPrintGridlines();
	public void setPrintGridlines(boolean enable);
	
	public int getHeaderMargin();
	public void setHeaderMargin(int px);
	public int getFooterMargin();
	public void setFooterMargin(int px);
	public int getLeftMargin();
	public void setLeftMargin(int px);
	public int getRightMargin();
	public void setRightMargin(int px);
	public int getTopMargin();
	public void setTopMargin(int px);
	public int getBottomMargin();
	public void setBottomMargin(int px);
	
	public void setPaperSize(PaperSize size);
	public PaperSize getPaperSize();
	
	public void setLandscape(boolean landscape);
	public boolean isLandscape();
	
	// TODO
//	public void setScale(short scale);
//	public short getScale();
	
}
