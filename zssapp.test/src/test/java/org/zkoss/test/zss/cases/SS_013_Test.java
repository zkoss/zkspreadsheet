/* SS_013_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 14, 2012 4:25:33 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import org.junit.Test;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_013_Test extends ZSSAppTest {
	
	@Test
	public void set_font_family_arial() {
		setFontFamilyAndVerify("arial", 11, 5, 16, 5);
	}
	
	@Test
	public void set_font_family_arial_black() {
		setFontFamilyAndVerify("arial-black", 11, 5, 16, 5);
	}
	
	@Test
	public void set_font_family_comic_sans_ms() {
		setFontFamilyAndVerify("comic-sans-ms", 11, 5, 16, 5);
	}
	
	@Test
	public void set_font_family_courier_new() {
		setFontFamilyAndVerify("courier-new", 11, 5, 16, 5);
	}
	
	@Test
	public void set_font_family_georgia() {
		setFontFamilyAndVerify("georgia", 11, 5, 16, 5);
	}
	
	@Test
	public void set_font_family_impact() {
		setFontFamilyAndVerify("impact", 11, 5, 16, 5);
	}
	
	@Test
	public void set_font_family_lucida_console() {
		setFontFamilyAndVerify("lucida-console", 11, 5, 16, 5);
	}
	
	@Test
	public void set_font_family_lucida_sans_unicode() {
		setFontFamilyAndVerify("lucida-sans-unicode", 11, 5, 16, 5);
	}
	
	@Test
	public void set_font_family_palatino_linotype() {
		setFontFamilyAndVerify("palatino-linotype", 11, 5, 16, 5);
	}
	
	@Test
	public void set_font_family_palatino_tahoma() {
		setFontFamilyAndVerify("tahoma", 11, 5, 16, 5);
	}
	
	@Test
	public void set_font_family_times_new_roman() {
		setFontFamilyAndVerify("times-new-roman", 11, 5, 16, 5);
	}
	
	@Test
	public void set_font_family_trebuchet_ms() {
		setFontFamilyAndVerify("trebuchet-ms", 11, 5, 16, 5);
	}
	
	@Test
	public void set_font_family_verdana() {
		setFontFamilyAndVerify("verdana", 11, 5, 16, 5);
	}
	
	@Test
	public void set_font_family_ms_sans_serif() {
		setFontFamilyAndVerify("ms-sans-serif", 11, 5, 16, 5);
	}
	
	@Test
	public void set_font_family_ms_serif() {
		setFontFamilyAndVerify("ms-serif", 11, 5, 16, 5);
	}
	
	void setFontFamilyAndVerify(String fontFamily, int tRow, int lCol, int bRow, int rCol) {
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		
		click(".zsfontfamily .z-combobox-btn");
		click(".zsfontfamily-" + fontFamily);
		
		verifyFontFamily(fontFamily.replace("-", " "), tRow, lCol, bRow, rCol);
	}
	

}
