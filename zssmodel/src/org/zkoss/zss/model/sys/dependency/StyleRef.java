package org.zkoss.zss.model.sys.dependency;

/**
 * This is a marking class to control whether to update cell's style or just
 * update cell's text. See RangeImpl#handleRefNotifyContentChange()
 *  
 * @author henrichen
 * @since 3.9.0
 */
public class StyleRef implements Ref {
	
	public static StyleRef inst = new StyleRef();

	@Override
	public RefType getType() {
		return RefType.STYLE;
	}

	@Override
	public String getBookName() {
		return null;
	}

	@Override
	public String getSheetName() {
		return null;
	}

	@Override
	public String getLastSheetName() {
		return null;
	}

	@Override
	public int getRow() {
		return 0;
	}

	@Override
	public int getColumn() {
		return 0;
	}

	@Override
	public int getLastRow() {
		return 0;
	}

	@Override
	public int getLastColumn() {
		return 0;
	}

	@Override
	public int getSheetIndex() {
		return 0;
	}

	@Override
	public int getLastSheetIndex() {
		return 0;
	}
}
