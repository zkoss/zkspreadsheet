package org.zkoss.zss.api;

import org.zkoss.zss.api.model.NSheet;
import org.zkoss.zss.ui.impl.Utils;

public class UnitUtil {

	
	
	
	/** convert pixel to EMU */
	public static int pxToEmu(int px) {
		//refer form ActionHandler
		return (int) Math.round(((double)px) * 72 * 20 * 635 / 96); //assume 96dpi
	}
	
	
	public static int getRowHeightInPx(NSheet sheet,int row){
		return Utils.getRowHeightInPx(sheet.getNative(), row);
	}
	
	public static int getColumnWidthInPx(NSheet sheet,int col){
		return Utils.getColumnWidthInPx(sheet.getNative(), col);
	}

	/**
	 * convert column width(char 256) to pixel
	 * @return the width in pixel
	 */
	public static int char256ToPx(int width256, int charWidthPx) {
		return Utils.fileChar256ToPx(width256,charWidthPx);
	}
	
	/**
	 * convert row height(twip) to pixel
	 * @return the height in pixel
	 */
	public static int twipToPx(int twip) {
		return Utils.twipToPx(twip);
	}
}
