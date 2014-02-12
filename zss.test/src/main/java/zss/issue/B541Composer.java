package zss.issue;

import org.zkoss.lang.Objects;
import org.zkoss.poi.ss.usermodel.RichTextString;
import org.zkoss.poi.ss.util.CellReference;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Button;
import org.zkoss.zul.Messagebox;

public class B541Composer extends SelectorComposer<Component>{

	@Wire
	Spreadsheet ss;
	
	
	public void autoFill(String area1,String area2,Button btn){
		Range r1 = Ranges.range(ss.getSelectedSheet(),area1);
		Range r2 = Ranges.range(ss.getSelectedSheet(),area2);
		r1.autoFill(r2, Range.FILL_DEFAULT);
		if(btn!=null){
			btn.setDisabled(true);
		}
	}
	
	public void checkValue(String area1,String area2,Button btn){
		Worksheet sheet = ss.getSelectedSheet();
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
				String editText1 = cell1.getEditText();
				String editText2 = cell2.getEditText();
				String display1 = getCellDisplay(cell1);
				String display2 = getCellDisplay(cell2);
				
				System.out.println("display text between "+a1.formatAsString()+"("+display1+") and "+a2.formatAsString()+"("+display2+")");
				System.out.println("edit text between "+a1.formatAsString()+"("+editText1+") and "+a2.formatAsString()+"("+editText1+")");
				
				if(!Objects.equals(display1, display2)){
					Messagebox.show("display text different between "+a1.formatAsString()+"("+display1+") and "+a2.formatAsString()+"("+display2+")");
					return;
				}
				if(!Objects.equals(editText1, editText2)){
					if(!compareDouble(editText1,editText2)){
						Messagebox.show("edit text different between "+a1.formatAsString()+"("+editText1+") and "+a2.formatAsString()+"("+editText2+")");
						return;
					}
				}
			}
		}
		if(btn!=null){
			btn.setDisabled(true);
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
		RichTextString text = range.getText();
		return text==null?null:text.getString();
	}
	
}
