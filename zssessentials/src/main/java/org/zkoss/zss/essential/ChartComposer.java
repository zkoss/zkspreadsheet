package org.zkoss.zss.essential;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.SheetAnchor;
import org.zkoss.zss.api.SheetOperationUtil;
import org.zkoss.zss.api.model.Chart;
import org.zkoss.zss.api.model.Chart.Grouping;
import org.zkoss.zss.api.model.Chart.LegendPosition;
import org.zkoss.zss.api.model.Chart.Type;
import org.zkoss.zss.api.model.ChartData;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zssex.api.ChartDataUtil;
import org.zkoss.zul.Intbox;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Listbox;

/**
 * Demonstrate chart related API usage
 * 
 * @author Hawk
 * 
 */
@SuppressWarnings("serial")
public class ChartComposer extends SelectorComposer<Component> {

	@Wire
	private Intbox toRowBox;
	@Wire
	private Intbox toColumnBox;
	@Wire
	private Spreadsheet ss;
	@Wire
	private Listbox chartListbox;

	private ListModelList<Chart> chartList = new ListModelList<Chart>();

	@Listen("onClick = #addButton")
	public void add() {
		Range selection = Ranges.range(ss.getSelectedSheet(),ss.getSelection());
		SheetAnchor selectionAnchor = SheetOperationUtil.toChartAnchor(selection);
		ChartData chartData = ChartDataUtil.getChartData(ss.getSelectedSheet(),
				new Rect("A1:B6"), Type.PIE);
		selection.addChart(selectionAnchor, chartData, Type.PIE, Grouping.STANDARD, LegendPosition.RIGHT);
		refreshChartList();
	}


	@Listen("onClick = #deleteButton")
	public void delete() {
		if (chartListbox.getSelectedItem() != null){
			Ranges.range(ss.getSelectedSheet())
				.deleteChart((Chart)chartListbox.getSelectedItem().getValue());
			refreshChartList();
		}
	}

	@Listen("onClick = #moveButton")
	public void move() {
		if (chartListbox.getSelectedItem() != null){
			//calculate destination anchor
			SheetAnchor fromAnchor = ((Chart) chartListbox.getSelectedItem()
					.getValue()).getAnchor();
			int rowOffset = fromAnchor.getLastRow() - fromAnchor.getRow();
			int columnOffset = fromAnchor.getLastColumn() - fromAnchor.getColumn();
			SheetAnchor toAnchor = new SheetAnchor(toRowBox.getValue(), 
					toColumnBox.getValue(),
					fromAnchor.getXOffset(), fromAnchor.getYOffset(),
					toRowBox.getValue()+rowOffset, toColumnBox.getValue()+columnOffset,
					fromAnchor.getLastXOffset(), fromAnchor.getLastYOffset());
			
			Ranges.range(ss.getSelectedSheet())
				.moveChart(toAnchor, (Chart)chartListbox.getSelectedItem().getValue());
			refreshChartList();
		}
	}

	private void refreshChartList(){
		chartList.clear();
		chartList.addAll(ss.getSelectedSheet().getCharts());
		chartListbox.setModel(chartList);
	}
	
	public void addWithUtil(){
		ChartData chartData = ChartDataUtil.getChartData(ss.getSelectedSheet(),ss.getSelection(), Type.PIE);
		SheetOperationUtil.addChart(Ranges.range(ss.getSelectedSheet(),ss.getSelection()),
		chartData, Type.PIE, Grouping.STANDARD, LegendPosition.RIGHT);
	}
}
