package org.zkoss.zss.issue;

import java.util.Locale;
import java.util.TimeZone;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.lang.Objects;
import org.zkoss.poi.ss.util.CellReference;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

/**
 * @author Hawk
 *
 */
public class Issue541Test {
	
	TimeZone oldTimeZone;
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssLocale(Locale.US);
		oldTimeZone = TimeZone.getDefault();
		TimeZone.setDefault(TimeZone.getTimeZone("America/New_York"));//timezone for issue 541
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssLocale();
		TimeZone.setDefault(oldTimeZone);
	}
	
	/**
	 * For auto fill:
	 * Notice that row count (column count) of destination range should be a multiple of source range.  
	 */
	@Test
	public void testFill(){
		Book book = Util.loadBook("541-daylightsave.xlsx");
		Sheet sheet = book.getSheetAt(0);
		autoFill(sheet,"A2","A2:A11");
		autoFill(sheet,"B12","B2:B12");
		autoFill(sheet,"C2:C3","C2:C11");
		autoFill(sheet,"D2:D3","D2:D11");
		autoFill(sheet,"E2","E2:E11");
		autoFill(sheet,"F2:F3","F2:F11");
		autoFill(sheet,"G2:G3","G2:G11");
		autoFill(sheet,"H2","H2:H11");
		autoFill(sheet,"I2:I3","I2:I11");
		autoFill(sheet,"J2:J3","J2:J11");
		autoFill(sheet,"K2:K3","K2:K11");
		autoFill(sheet,"L2:L3","L2:L11");
		
		checkValue(sheet,"A2:A11","A15:A24");
		checkValue(sheet,"B2:B12","B15:B25");
		checkValue(sheet,"C2:C11","C15:C24");
		checkValue(sheet,"D2:D11","D15:D24");
		checkValue(sheet,"E2:E11","E15:E24");
		checkValue(sheet,"F2:F11","F15:F24");
		checkValue(sheet,"G2:G11","G15:G24");
		checkValue(sheet,"H2:H11","H15:H24");
		checkValue(sheet,"I2:I11","I15:I24");
		checkValue(sheet,"J2:J11","J15:J24");
		checkValue(sheet,"K2:K11","K15:K24");
		checkValue(sheet,"L2:L11","L15:L24");
		
	}
	
	public void autoFill(Sheet sheet,String area1,String area2){
		Range r1 = Ranges.range(sheet,area1);
		Range r2 = Ranges.range(sheet,area2);
		r1.autoFill(r2, Range.AutoFillType.DEFAULT);
	}
	
	private void checkValue(Sheet sheet,String area1,String area2){
		Range r1 = Ranges.range(sheet,area1);
		Range r2 = Ranges.range(sheet,area2);
		int rsize = Math.min(r1.getLastRow()-r1.getRow()+1, r2.getLastRow()-r2.getRow()+1);
		int csize = Math.min(r1.getLastColumn()-r1.getColumn()+1,  r2.getLastColumn()-r2.getColumn()+1);
		
		for(int i=0;i<rsize;i++){
			for(int j=0;j<csize;j++){
				CellReference a1 = new CellReference(r1.getRow()+i,r1.getColumn()+j);
				CellReference a2 = new CellReference(r2.getRow()+i,r2.getColumn()+j);
				Range cell1 = Ranges.range(sheet,a1.formatAsString());
				Range cell2 = Ranges.range(sheet,a2.formatAsString());
				String editText1 = cell1.getCellEditText();
				String editText2 = cell2.getCellEditText();
				String display1 = getCellDisplay(cell1);
				String display2 = getCellDisplay(cell2);
				
				System.out.println("display text between "+a1.formatAsString()+"("+display1+") and "+a2.formatAsString()+"("+display2+")");
				System.out.println("edit text between "+a1.formatAsString()+"("+editText1+") and "+a2.formatAsString()+"("+editText2+")");
				
				if(!Objects.equals(display1, display2)){
					Assert.fail("display text different between "+a1.formatAsString()+"("+display1+") and "+a2.formatAsString()+"("+display2+")");
					return;
				}
				if(!Objects.equals(editText1, editText2)){
					if(!compareDouble(editText1,editText2)){
						Assert.fail("edit text different between "+a1.formatAsString()+"("+editText1+") and "+a2.formatAsString()+"("+editText2+")");
						return;
					}
				}
			}
		}
	}
	private boolean compareDouble(String editText1, String editText2) {
		return compareDouble(editText1,editText2,0.00000001);
	}
	private boolean compareDouble(String editText1, String editText2, double eplision) {
		try{
			double d1 = Double.parseDouble(editText1);
			double d2 = Double.parseDouble(editText2);
			return d1==d2? true:Math.abs(d1-d2) < eplision;
		}catch(Exception x){
			return false;
		}
	}

	private String getCellDisplay(Range range) {
		String text = range.getCellFormatText();
		return text;
	}
	
	@Test
	public void pasteAutoRepeat(){
		Book book = Util.loadBook("blank.xlsx");
		book.getSheetAt(0);
	}
}
