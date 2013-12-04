package org.zkoss.zss.ngmodel;

public interface NColor {

	/**
	 * Gets the color html string expression, for example "#FF6622"
	 * @return
	 */
	public String getHtmlColor();
	
	/**
	 * Gets the color types array, order by Red,Green,Blue
	 * @return
	 */
	public byte[] getRGB();
	
	
}
