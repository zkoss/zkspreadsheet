package zss.issue;

import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.*;
import org.zkoss.zss.api.*;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Vlayout;

import java.util.*;

public class OverflowRightAlignComposer extends SelectorComposer<Vlayout> {

    @Wire
    private Spreadsheet ss;
    private LinkedList<Range> rangesToNotify = new LinkedList<Range>();

    @Listen("onClick = #fill")
    public void fill() {
        for (int i= 0 ; i < 50 ; i++) {
            getRangeWithoutRefresh(i, 0, i, 9).setCellValue(new Random().nextInt(100));
        }
        notifyRanges();
    }

    @Listen("onClick = #unfreeze")
    public void unfreeze(){
        Ranges.range(ss.getSelectedSheet()).setFreezePanel(0, 0);
    }

    @Listen("onClick = #freeze")
    public void freeze(){
        Ranges.range(ss.getSelectedSheet()).setFreezePanel(0, 3);
    }

    private void notifyRanges() {
        Ranges.range(ss.getSelectedSheet()).notifyChange();
//        while (!rangesToNotify.isEmpty()){
//            rangesToNotify.poll().notifyChange();
//        }
        rangesToNotify.clear();
    }

    private Range getRangeWithoutRefresh(int topRow, int leftColumn, int bottomRow, int rightColumn){
        Range range = Ranges.range(ss.getSelectedSheet(), topRow, leftColumn, bottomRow, rightColumn);
        range.setAutoRefresh(false);
        rangesToNotify.add(range);
        return range;
    }
}
