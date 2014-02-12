package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;

import java.io.UnsupportedEncodingException;
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
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zul.Button;
import org.zkoss.zul.Messagebox;

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
		Setup.pushZssContextLocale(Locale.US);
		oldTimeZone = TimeZone.getDefault();
		TimeZone.setDefault(TimeZone.getTimeZone("America/New_York"));//timezone for issue 541
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssContextLocale();
		TimeZone.setDefault(oldTimeZone);
	}
	
	
	@Test
	public void testFill(){
		Book book = Util.loadBook(this,"book/541-daylightsave.xlsx");
		Sheet sheet = book.getSheetAt(0);
		autoFill(sheet,"A2","A2:A12");
		autoFill(sheet,"B12","B2:B12");
		autoFill(sheet,"C2:C3","C2:C12");
		autoFill(sheet,"D2:D3","D2:D12");
		autoFill(sheet,"E2","E2:E12");
		autoFill(sheet,"F2:F3","F2:F12");
		autoFill(sheet,"G2:G3","G2:G12");
		autoFill(sheet,"H2","H2:H12");
		autoFill(sheet,"I2:I3","I2:I12");
		autoFill(sheet,"J2:J3","J2:J12");
		autoFill(sheet,"K2:K3","K2:K12");
		autoFill(sheet,"L2:L3","L2:L12");
		
		checkValue(sheet,"A2:A12","A15:A25");
		checkValue(sheet,"B2:B12","B15:B25");
		checkValue(sheet,"C2:C12","C15:C25");
		checkValue(sheet,"D2:D12","D15:D25");
		checkValue(sheet,"E2:E12","E15:E25");
		checkValue(sheet,"F2:F12","F15:F25");
		checkValue(sheet,"G2:G12","G15:G25");
		checkValue(sheet,"H2:H12","H15:H25");
		checkValue(sheet,"I2:I12","I15:I25");
		checkValue(sheet,"J2:J12","J15:J25");
		checkValue(sheet,"K2:K12","K15:K25");
		checkValue(sheet,"L2:L12","L15:L25");
		
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
	
}
