package zss.issue;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.SortDataOption;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.ui.Spreadsheet;

public class Composer1262SortPerformance extends SelectorComposer<Component> {
    private static final long serialVersionUID = 7216068332971922532L;
    @Wire
    Spreadsheet ss;
    
    
    @Listen("onClick = #sort")
    public void onCellClick(Event event){
    	int col = ss.getSelection().getColumn();
    	Range sortingRange = Ranges.range(ss.getSelectedSheet(), 1, 0, 165, 33);
    	Range index = Ranges.range(ss.getSelectedSheet(), 1, col, 165, col);
    	//10th hasHeader, means the sorting range contains a header data (not for sorting, e.g. Country), since sortingRange starts from 1, it doesn't contains a header
    	sortingRange.sort(index, false, SortDataOption.NORMAL_DEFAULT, null, true, SortDataOption.NORMAL_DEFAULT, null, true, SortDataOption.NORMAL_DEFAULT, false, true, false);
    }
    
    @Listen("onClick = #sort2")
    public void sort(Event event){
        int col = ss.getSelection().getColumn();
        Range index = Ranges.range(ss.getSelectedSheet(), 1, col, 165, col);
    	Range sortingRange = Ranges.range(ss.getSelectedSheet(), 1, 0, 165, 33);
        sortingRange.setAutoRefresh(false);
        sortingRange.sort(index, false, SortDataOption.NORMAL_DEFAULT, null, true, SortDataOption.NORMAL_DEFAULT, null, true, SortDataOption.NORMAL_DEFAULT, false, true, false);
        sortingRange.notifyChange();
    }
}
