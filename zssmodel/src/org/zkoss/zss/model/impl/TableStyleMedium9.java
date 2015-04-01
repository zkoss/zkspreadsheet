/* TableStyleMedium9.java

	Purpose:
		
	Description:
		
	History:
		Mar 30, 2015 5:17:51 PM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import org.zkoss.zss.model.SBorder.BorderType;
import org.zkoss.zss.model.SFill.FillPattern;
import org.zkoss.zss.model.STableStyleElem;
import org.zkoss.zss.model.SFont.Underline;

/**
 * Builtin TableStyleMediumn9
 * 
 * @author henri
 * @since 3.8.0
 */
public class TableStyleMedium9 extends TableStyleImpl {
	private TableStyleMedium9() {
		super(
				"TableStyleMedium9",//name,
				M9_Whole_Table, 	//wholeTable,
				M9_Col_Stripe1, 	//colStripe1,
				1, 					//colStripe1Size,
				null, 				//colStripe2,
				1, 					//colStripe2Size,
				M9_Row_Stripe1, 	//rowStripe1,
				1, 					//rowStripe1Size,
				null,				//rowStripe2,
				1, 					//rowStripe2Size,
				M9_Last_Col, 	//lastCol,
				M9_First_Col,	//firstCol,
				M9_Header_Row,	//headerRow,
				M9_Total_Row,	//totalRow,
				null,	//firstHeaderCell,
				null,	//lastHeaderCell,
				null,	//firstTotalCell,
				null	//lastTotalCell
			);
	}
	
	// TableStyleMedium9
	private static final STableStyleElem M9_Col_Stripe1 = 
			new TableStyleElemImpl(
				null,	//font
				new FillImpl(FillPattern.SOLID, "B8CCE4", "B8CCE4"), 	//fill
				null	//border
			);

	private static final STableStyleElem M9_Row_Stripe1 =
			M9_Col_Stripe1;
	private static final STableStyleElem M9_Last_Col =
			new TableStyleElemImpl(
				new FontImpl("FFFFFF", true /*bold*/, false /*fontItalic*/, 
					false /*fontStrikeout*/, Underline.NONE), 			//font
				new FillImpl(FillPattern.SOLID, "4F81BD", "4F81BD"),	//fill
				null //border
			);
	private static final STableStyleElem M9_First_Col =
			M9_Last_Col;
	private static final STableStyleElem M9_Total_Row =
			new TableStyleElemImpl(
					new FontImpl("FFFFFF", true /*bold*/, false /*fontItalic*/, 
							false /*fontStrikeout*/, Underline.NONE), 	//font
					new FillImpl(FillPattern.SOLID, "4F81BD", "4F81BD"),//fill
					new BorderImpl(
							null, //left 
							new BorderLineImpl(BorderType.THICK, "FFFFFF"), //top 
							null, //right
							null, //bottom
							null, //diagonal
							null, //vertical
							null  //horizontal
					)
			);
	private static final STableStyleElem M9_Header_Row =
			new TableStyleElemImpl(
					new FontImpl("FFFFFF", true /*bold*/, false /*fontItalic*/, 
							false /*fontStrikeout*/, Underline.NONE), 	//font
					new FillImpl(FillPattern.SOLID, "4F81BD", "4F81BD"),//fill
					new BorderImpl(
							null, //left 
							null, //top
							null, //right
							new BorderLineImpl(BorderType.THICK, "FFFFFF"), //bottom 
							null, //diagonal
							null, //vertical
							null  //horizontal
					)
			);
	private static final STableStyleElem M9_Whole_Table = 
			new TableStyleElemImpl(
					new FontImpl("000000", false /*bold*/, false /*fontItalic*/, 
							false /*fontStrikeout*/, Underline.NONE), 	//font
					new FillImpl(FillPattern.SOLID, "DBE5F1", "DBE5F1"),//fill
					new BorderImpl(
							null, //left 
							null, //top 
							null, //right
							null, //bottom
							null, //diagonal
							new BorderLineImpl(BorderType.THIN, "FFFFFF"),	//vertical
							new BorderLineImpl(BorderType.THIN, "FFFFFF") 	//horizontal
					)
			);

	public static final TableStyleMedium9 instance = new TableStyleMedium9();
}
