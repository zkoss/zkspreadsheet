/* DateInputMaskTest.java

	Purpose:
		
	Description:
		
	History:
		Mar 15, 2010 3:31:34 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.impl;


import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.assertEquals;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

/**
 * Test date input mask.
 * @author henrichen
 *
 */
public class DateInputMaskTest {
	DateInputMask _formater;

	private static final String OK[][] = {
		
		{"Feb-1, 2", "d-mmm-yy"},
		{"Feb -1, 2", "d-mmm-yy"},
		{"Feb- 1, 2", "d-mmm-yy"},
		{"Feb - 1, 2", "d-mmm-yy"},
		{"Feb-1, 20", "d-mmm-yy"},
		{"Feb -1, 20", "d-mmm-yy"},
		{"Feb- 1, 20", "d-mmm-yy"},
		{"Feb - 1, 20", "d-mmm-yy"},
		{"Feb-1, 2000", "d-mmm-yy"},
		{"Feb -1, 2000", "d-mmm-yy"},
		{"Feb- 1, 2000", "d-mmm-yy"},
		{"Feb - 1, 2000", "d-mmm-yy"},
		{"Feb 1, 2", "d-mmm-yy"},
		{"Feb  1, 2", "d-mmm-yy"},
		{"Feb  1, 2", "d-mmm-yy"},
		{"Feb   1, 2", "d-mmm-yy"},
		{"Feb-1 , 2", "d-mmm-yy"},
		{"Feb -1 , 2", "d-mmm-yy"},
		{"Feb- 1 , 2", "d-mmm-yy"},
		{"Feb - 1 , 2", "d-mmm-yy"},

		{"Feb 1 , 2", "d-mmm-yy"},
		{"Feb  1 , 2", "d-mmm-yy"},
		{"Feb  1 , 2", "d-mmm-yy"},
		{"Feb   1 , 2", "d-mmm-yy"},
		
		{"Feb-1", "d-mmm"},
		{"Feb-1", "d-mmm"},
		{"Feb -1", "d-mmm"}, 
		{"Feb -1", "d-mmm"},
		{"Feb- 1",  "d-mmm"},
		{"Feb- 1", "d-mmm"},
		{"Feb - 1",  "d-mmm"},
		{"Feb - 1", "d-mmm"},
		
		{"Feb 1", "d-mmm"},
		{"Feb 1",  "d-mmm"}, 
		{"Feb  1", "d-mmm"},
		{"Feb  1",  "d-mmm"},
		{"Feb  1", "d-mmm"}, 
		{"Feb  1", "d-mmm"},
		{"Feb   1",  "d-mmm"},
		{"Feb   1", "d-mmm"},
		
		{"Feb-1", "d-mmm"},  
		{"Feb-1",  "d-mmm"},
		{"Feb -1",   "d-mmm"},
		{"Feb -1",  "d-mmm"},
		{"Feb- 1",   "d-mmm"},
		{"Feb- 1",  "d-mmm"},
		{"Feb - 1",   "d-mmm"},
		{"Feb - 1",  "d-mmm"},
		
		{"Feb 1", "d-mmm"},  
		{"Feb 1",  "d-mmm"},
		{"Feb  1",   "d-mmm"},
		{"Feb  1",  "d-mmm"},
		{"Feb  1",   "d-mmm"},
		{"Feb  1",  "d-mmm"},
		{"Feb   1",   "d-mmm"},
		{"Feb   1",  "d-mmm"},
		
		{"2-1", "d-mmm"},
		{"2 -1", "d-mmm"},
		{"2- 1", "d-mmm"},
		{"2 - 1", "d-mmm"},

		{"2/1", "d-mmm"},
		{"2 /1", "d-mmm"},
		{"2/ 1", "d-mmm"},
		{"2 / 1", "d-mmm"},

		{"2/1/3", "m/d/yyyy"},
		{"2 /1/3", "m/d/yyyy"},
		{"2/ 1/3", "m/d/yyyy"},
		{"2 / 1/3", "m/d/yyyy"},

		{"2/1 /3", "m/d/yyyy"},
		{"2/1/ 3", "m/d/yyyy"},
		{"2/1 / 3", "m/d/yyyy"},

		{"2 /1 /3", "m/d/yyyy"},
		{"2 /1/ 3", "m/d/yyyy"},
		{"2 /1 / 3", "m/d/yyyy"},

		{"2/ 1 /3", "m/d/yyyy"},
		{"2/ 1/ 3", "m/d/yyyy"},
		{"2/ 1 / 3", "m/d/yyyy"},

		{"2 / 1 /3", "m/d/yyyy"},
		{"2 / 1/ 3", "m/d/yyyy"},
		{"2 / 1 / 3", "m/d/yyyy"},
		
		{"2-1-3", "m/d/yyyy"},
		{"2 -1-3", "m/d/yyyy"},
		{"2- 1-3", "m/d/yyyy"},
		{"2 - 1-3", "m/d/yyyy"},

		{"2-1 -3", "m/d/yyyy"},
		{"2-1- 3", "m/d/yyyy"},
		{"2-1 - 3", "m/d/yyyy"},

		{"2 -1 -3", "m/d/yyyy"},
		{"2 -1- 3", "m/d/yyyy"},
		{"2 -1 - 3", "m/d/yyyy"},

		{"2- 1 -3", "m/d/yyyy"},
		{"2- 1- 3", "m/d/yyyy"},
		{"2- 1 - 3", "m/d/yyyy"},

		{"2 - 1 -3", "m/d/yyyy"},
		{"2 - 1- 3", "m/d/yyyy"},
		{"2 - 1 - 3", "m/d/yyyy"},
		
		{"2/May/3", "d-mmm-yy"},
		{"2 /May/3", "d-mmm-yy"},
		{"2/ May/3", "d-mmm-yy"},
		{"2 / May/3", "d-mmm-yy"},

		{"2/May /3", "d-mmm-yy"},
		{"2/May/ 3", "d-mmm-yy"},
		{"2/May / 3", "d-mmm-yy"},

		{"2 /May /3", "d-mmm-yy"},
		{"2 /May/ 3", "d-mmm-yy"},
		{"2 /May / 3", "d-mmm-yy"},

		{"2/ May /3", "d-mmm-yy"},
		{"2/ May/ 3", "d-mmm-yy"},
		{"2/ May / 3", "d-mmm-yy"},

		{"2 / May /3", "d-mmm-yy"},
		{"2 / May/ 3", "d-mmm-yy"},
		{"2 / May / 3", "d-mmm-yy"},
		
		{"2-May-3", "d-mmm-yy"},
		{"2 -May-3", "d-mmm-yy"},
		{"2- May-3", "d-mmm-yy"},
		{"2 - May-3", "d-mmm-yy"},

		{"2-May -3", "d-mmm-yy"},
		{"2-May- 3", "d-mmm-yy"},
		{"2-May - 3", "d-mmm-yy"},

		{"2 -May -3", "d-mmm-yy"},
		{"2 -May- 3", "d-mmm-yy"},
		{"2 -May - 3", "d-mmm-yy"},

		{"2- May -3", "d-mmm-yy"},
		{"2- May- 3", "d-mmm-yy"},
		{"2- May - 3", "d-mmm-yy"},

		{"2 - May -3", "d-mmm-yy"},
		{"2 - May- 3", "d-mmm-yy"},
		{"2 - May - 3", "d-mmm-yy"},

		{"2-30", "mmm-yy"},
		{"2- 30", "mmm-yy"},
		{"2 -30", "mmm-yy"},
		{"2 - 30", "mmm-yy"},
		{"Feb-30", "mmm-yy"},
		{"Feb- 30", "mmm-yy"},
		{"Feb -30", "mmm-yy"},
		{"Feb - 30", "mmm-yy"},
		{"Feb/30", "mmm-yy"},
		{"Feb/ 30", "mmm-yy"},
		{"Feb /30", "mmm-yy"},
		{"Feb / 30", "mmm-yy"},
		{"Feb 30", "mmm-yy"},
		{"Feb  30", "mmm-yy"},

		{"2-2030", "mmm-yy"},
		{"2- 2030", "mmm-yy"},
		{"2 -2030", "mmm-yy"},
		{"2 - 2030", "mmm-yy"},
		{"Feb-2030", "mmm-yy"},
		{"Feb- 2030", "mmm-yy"},
		{"Feb -2030", "mmm-yy"},
		{"Feb - 2030", "mmm-yy"},
		{"Feb/2030", "mmm-yy"},
		{"Feb/ 2030", "mmm-yy"},
		{"Feb /2030", "mmm-yy"},
		{"Feb / 2030", "mmm-yy"},
		{"Feb 2030", "mmm-yy"},
		{"Feb  2030", "mmm-yy"},
		
		{"10:2:1 p", "h:mm:ss AM/PM"},
		{"10 :2:1 p", "h:mm:ss AM/PM"},
		{"10: 2:1 p", "h:mm:ss AM/PM"},
		{"10 : 2:1 p", "h:mm:ss AM/PM"},

		{"10:2 :1 p", "h:mm:ss AM/PM"},
		{"10:2: 1 p", "h:mm:ss AM/PM"},
		{"10:2 : 1 p", "h:mm:ss AM/PM"},

		{"10 :2 :1 p", "h:mm:ss AM/PM"},
		{"10 :2: 1 p", "h:mm:ss AM/PM"},
		{"10 :2 : 1 p", "h:mm:ss AM/PM"},

		{"10: 2 :1 p", "h:mm:ss AM/PM"},
		{"10: 2: 1 p", "h:mm:ss AM/PM"},
		{"10: 2 : 1 p", "h:mm:ss AM/PM"},

		{"10 : 2 :1 p", "h:mm:ss AM/PM"},
		{"10 : 2: 1 p", "h:mm:ss AM/PM"},
		{"10 : 2 : 1 p", "h:mm:ss AM/PM"},

		{"10:2:1", "h:mm:ss"}, 
		{"10 :2:1", "h:mm:ss"},
		{"10: 2:1", "h:mm:ss"},
		{"10 : 2:1", "h:mm:ss"},  

		{"10:2 :1", "h:mm:ss"}, 
		{"10:2: 1", "h:mm:ss"}, 
		{"10:2 : 1", "h:mm:ss"},

		{"10 :2 :1", "h:mm:ss"}, 
		{"10 :2: 1", "h:mm:ss"}, 
		{"10 :2 : 1", "h:mm:ss"},

		{"10: 2 :1", "h:mm:ss"}, 
		{"10: 2: 1", "h:mm:ss"}, 
		{"10: 2 : 1", "h:mm:ss"},

		{"10 : 2 :1", "h:mm:ss"},
		{"10 : 2: 1", "h:mm:ss"}, 
		{"10 : 2 : 1", "h:mm:ss"},
		
		{"10:1:1.456 ", "mm:ss.0"},
		{"10 :1:1.456", "mm:ss.0"},
		{"10: 1:1.456", "mm:ss.0"},
		{"10 : 1:1.456  ", "mm:ss.0"},

		{"10:1 :1.45", "mm:ss.0"},
		{"10:1: 1.45 ", "mm:ss.0"},
		{"10:1 : 1.45", "mm:ss.0"},

		{"10 :1 :1.4 ", "mm:ss.0"},
		{"10 :1: 1.4 ", "mm:ss.0"},
		{"10 :1 : 1.4", "mm:ss.0"},

		{"10: 1 :1.4567", "mm:ss.0"}, 
		{"10: 1: 1.4567 ", "mm:ss.0"},
		{"10: 1 : 1.4567", "mm:ss.0"},
		
		{"10:", "h:mm"},
		{"10:20:", "h:mm"},
		{"10:20", "h:mm"},
		{"10:20.", "h:mm"},
		
		{"10: am", "h:mm AM/PM"},
		{"10:20: a", "h:mm AM/PM"},
		{"10:20 p", "h:mm AM/PM"},
		{"10:20. p", "h:mm AM/PM"},
		
		{"2-May -3 10:20.", "m/d/yyyy h:mm"},
		{"2-May- 3 10:20.", "m/d/yyyy h:mm"},
		{"2-May - 3 10:20.", "m/d/yyyy h:mm"},

		{"2 -May -3 10:20.", "m/d/yyyy h:mm"},
		{"2 -May- 3 10:20.", "m/d/yyyy h:mm"},
		{"2 -May - 3 10:20.", "m/d/yyyy h:mm"},

		{"2- May -3 10 :1 :1.4", "mm:ss.0"},
		{"2- May- 3 10 :1 :1.4", "mm:ss.0"},
		{"2- May - 3 10 :1 :1.4", "mm:ss.0"},

		{"2 - May -3 10 :1 :1.4", "mm:ss.0"},
		{"2 - May- 3 10 :1 :1.4", "mm:ss.0"},
		{"2 - May - 3 10 :1 :1.4", "mm:ss.0"},

		{"2-30 10:2:1 p", "m/d/yyyy h:mm"},
		{"2- 30 10:2:1 p", "m/d/yyyy h:mm"},
		{"2 -30 10:2:1 p", "m/d/yyyy h:mm"},
		{"2 - 30 10:2:1 p", "m/d/yyyy h:mm"},
		{"Feb-30 10:2:1 p", "m/d/yyyy h:mm"},
		{"Feb- 30 10:2:1 p", "m/d/yyyy h:mm"},
		{"Feb -30 10:2:1 p", "m/d/yyyy h:mm"},
		{"Feb - 30 10:2:1 p", "m/d/yyyy h:mm"},
		{"Feb/30 10:2:1 p", "m/d/yyyy h:mm"},
		{"Feb/ 30 10:2:1 p", "m/d/yyyy h:mm"},
		{"Feb /30 10:2:1 p", "m/d/yyyy h:mm"},
		{"Feb / 30 10:2:1 p", "m/d/yyyy h:mm"},
		{"Feb 30 10:2:1 p", "m/d/yyyy h:mm"},
		{"Feb  30 10:2:1 p", "m/d/yyyy h:mm"},

		{"2-2030 10:2:1 p", "m/d/yyyy h:mm"},
		{"2- 2030 10:2:1 p", "m/d/yyyy h:mm"},
		{"2 -2030 10:2:1 p", "m/d/yyyy h:mm"},
		{"2 - 2030 10:2:1 p", "m/d/yyyy h:mm"},
		{"Feb-2030 10:2:1 p", "m/d/yyyy h:mm"},
		{"Feb- 2030 10:2:1 p", "m/d/yyyy h:mm"},
		{"Feb -2030 10:2:1 p", "m/d/yyyy h:mm"},
		{"Feb - 2030 10:2:1 p", "m/d/yyyy h:mm"},
		{"Feb/2030 10:2:1 p", "m/d/yyyy h:mm"},
		{"Feb/ 2030 10:2:1 p", "m/d/yyyy h:mm"},
		{"Feb /2030 10:2:1 p", "m/d/yyyy h:mm"},
		{"Feb / 2030 10:2:1 p", "m/d/yyyy h:mm"},
		{"Feb 2030 10:2:1 p", "m/d/yyyy h:mm"},
		{"Feb  2030 10:2:1 p", "m/d/yyyy h:mm"},
	};
	
	private static final String FAIL[] = {
		"Feb-1,2",
		"Feb -1,2",
		"Feb- 1,2",
		"Feb - 1,2",
		"Feb-1,20",
		"Feb -1,20",
		"Feb- 1,20",
		"Feb - 1,20",
		"Feb-1, 200",
		"Feb-1,200",
		"Feb -1,200",
		"Feb- 1,200",
		"Feb - 1,200",
		"Feb-1,2000",
		"Feb -1,2000",
		"Feb- 1,2000",
		"Feb - 1,2000",
		"Feb-1, ",
		"Feb-1,",
		"Feb -1, ",
		"Feb -1,",
		"Feb- 1, ",
		"Feb- 1,",
		"Feb - 1, ",
		"Feb - 1,",
		
		"Feb 1,2",
		"Feb  1,2",
		"Feb  1,2",
		"Feb   1,2",
		"Feb 1, ",
		"Feb 1,",
		"Feb  1, ",
		"Feb  1,",
		"Feb  1, ",
		"Feb  1,",
		"Feb   1, ",
		"Feb   1,",
		
		"Feb-1 ,2",
		"Feb -1 ,2",
		"Feb- 1 ,2",
		"Feb - 1 ,2",
		"Feb-1 , ",
		"Feb-1 ,",
		"Feb -1 , ",
		"Feb -1 ,",
		"Feb- 1 , ",
		"Feb- 1 ,",
		"Feb - 1 , ",
		"Feb - 1 ,",
		
		"Feb 1 ,2",
		"Feb  1 ,2",
		"Feb  1 ,2",
		"Feb   1 ,2",
		"Feb 1 , ",
		"Feb 1 ,",
		"Feb  1 , ",
		"Feb  1 ,",
		"Feb  1 , ",
		"Feb  1 ,",
		"Feb   1 , ",
		"Feb   1 ,",

		"Feb/1/3",
		"Feb /1/3",
		"Feb/ 1/3",
		"Feb / 1/3",

		"Feb/1 /3",
		"Feb/1/ 3",
		"Feb/1 / 3",

		"Feb /1 /3",
		"Feb /1/ 3",
		"Feb /1 / 3",

		"Feb/ 1 /3",
		"Feb/ 1/ 3",
		"Feb/ 1 / 3",

		"Feb / 1 /3",
		"Feb / 1/ 3",
		"Feb / 1 / 3",
		
		"Feb-1-3",
		"Feb -1-3",
		"Feb- 1-3",
		"Feb - 1-3",

		"Feb-1 -3",
		"Feb-1- 3",
		"Feb-1 - 3",

		"Feb -1 -3",
		"Feb -1- 3",
		"Feb -1 - 3",

		"Feb- 1 -3",
		"Feb- 1- 3",
		"Feb- 1 - 3",

		"Feb - 1 -3",
		"Feb - 1- 3",
		"Feb - 1 - 3",
		
		"2-1899",
		"2- 1899",
		"2 -1899",
		"2 - 1899",
		"Feb-1899",
		"Feb- 1899",
		"Feb -1899",
		"Feb - 1899",
		"Feb/1899",
		"Feb/ 1899",
		"Feb /1899",
		"Feb / 1899",
		"Feb 1899",
		"Feb  1899",
		
		"10 : 1 :1.45678 a",
		"10 : 1: 1.45678 ",
		"10 : 1 : 1.45678 ",
		"10::",
		"13 a",

	};
	
	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		_formater = new DateInputMask();
	}

	/**
	 * @throws java.lang.Exception
	 */
	@After
	public void tearDown() throws Exception {
		_formater = null;
	}

	@Test
	public void testParseDateInput() {
		for(int j= 0; j < OK.length; ++j) {
			Object[] result = _formater.parseDateInput(OK[j][0]);
			assertTrue(OK[j][0]+"->"+j,  result[1] != null);
			assertEquals(OK[j][0]+"->"+j, OK[j][1], (String) result[1]);
		}
		for(int j= 0; j < FAIL.length; ++j) {
			Object[] result = _formater.parseDateInput(FAIL[j]);
			assertFalse(FAIL[j]+"->"+j, result[1] != null);
			assertEquals(FAIL[j]+"->"+j, FAIL[j], (String)result[0]);
		}
	}
}
