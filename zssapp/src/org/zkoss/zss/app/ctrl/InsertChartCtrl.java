/* InsertChartCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 15, 2011 9:37:25 AM , Created by sam
}}IS_NOTE

Copyright (C) 2011 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.app.ctrl;

import java.util.Map;

//import org.zkoss.poi.ss.usermodel.charts.ChartType;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.api.model.Chart;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.Dropdownbutton;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.app.zul.Zssapp;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zss.app.zul.ctrl.WorkbookCtrl;
import org.zkoss.zul.Menuitem;

/**
 * @author sam
 *
 */
public class InsertChartCtrl extends GenericForwardComposer {
	
	/*views*/
	private Dropdownbutton insertChartBtn;
	private Menuitem insertColumnChart;
	private Menuitem insertColumnChart3D;
	private Menuitem insertBarChart;
	private Menuitem insertBarChart3D;
	private Menuitem insertLineChart;
	private Menuitem insertLineChart3D;
	private Menuitem insertPieChart;
	private Menuitem insertPieChart3D;
//	private Menuitem insertOfPieChart; //Not support yet
	private Menuitem insertAreaChart;
//	private Menuitem insertAreaChart3D; //Not support yet
//	private Menuitem insertSurfaceChart; //Not support yet
//	private Menuitem insertSurfaceChart3D; //Not support yet
//	private Menuitem insertBubbleChart; //Not support yet
	private Menuitem insertDoughnutChart;
//	private Menuitem insertRadarChart; //Not support yet
	private Menuitem insertScatterChart;
//	private Menuitem insertStockChart; //Not support yet
	
	private Dialog insertChartAtDialog;
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		boolean disable = !WebApps.getFeature("pe");
		insertColumnChart.setDisabled(disable);
		insertColumnChart3D.setDisabled(disable);
		insertBarChart.setDisabled(disable);
		insertBarChart3D.setDisabled(disable);
		insertLineChart.setDisabled(disable);
		insertLineChart3D.setDisabled(disable);
		insertPieChart.setDisabled(disable);
		insertPieChart3D.setDisabled(disable);
		insertAreaChart.setDisabled(disable);
		insertDoughnutChart.setDisabled(disable);
		insertScatterChart.setDisabled(disable);
	}

	public void onClick$insertColumnChart() {
		insertChart(Chart.Type.COLUMN);
	}
	
	public void onClick$insertColumnChart3D() {
		insertChart(Chart.Type.COLUMN_3D);
	}
	
	public void onClick$insertBarChart() {
		insertChart(Chart.Type.BAR);
	}
	
	public void onClick$insertBarChart3D() {
		insertChart(Chart.Type.BAR_3D);
	}
	
	public void onClick$insertLineChart() {
		insertChart(Chart.Type.LINE);
	}
	
	public void onClick$insertLineChart3D() {
		insertChart(Chart.Type.LINE_3D);
	}
	
	public void onClick$insertPieChart() {
		insertChart(Chart.Type.PIE);
	}
	
	public void onClick$insertPieChart3D() {
		insertChart(Chart.Type.PIE_3D);
	}
	
	public void onClick$insertOfPieChart() {
		insertChart(Chart.Type.OF_PIE);
	}
	
	public void onClick$insertAreaChart() {
		insertChart(Chart.Type.AREA);
	}
	
	public void onClick$insertAreaChart3D() {
		insertChart(Chart.Type.AREA_3D);
	}
	
	public void onClick$insertSurfaceChart() {
		insertChart(Chart.Type.SURFACE);
	}
	
	public void onClick$insertSurfaceChart3D() {
		insertChart(Chart.Type.SURFACE_3D);
	}
	
	public void onClick$insertBubbleChart() {
		insertChart(Chart.Type.BUBBLE);
	}
	
	public void onClick$insertDoughnutChart() {
		insertChart(Chart.Type.DOUGHNUT);
	}
	
	public void onClick$insertRadarChart() {
		insertChart(Chart.Type.RADAR);
	}
	
	public void onClick$insertScatterChart() {
		insertChart(Chart.Type.SCATTER);
	}
	
	public void onClick$insertStockChart() {
		insertChart(Chart.Type.STOCK);
	}
	
	public void onDropdown$insertChartBtn() {
		getWorkbookCtrl().reGainFocus();
	}
	
	protected WorkbookCtrl getWorkbookCtrl() {
		return getDesktopWorkbenchContext().getWorkbookCtrl();
	}
	
	protected DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return Zssapp.getDesktopWorkbenchContext(self);
	}
	
	protected void insertChart(Chart.Type chartType) {
		//open dialog to get insert at
		if (insertChartAtDialog == null) {
			insertChartAtDialog = (Dialog)Executions.createComponents(Consts._InsertWidgetAtDialog_zul, (Component)spaceOwner, null);
			insertChartAtDialog.addEventListener("onClose", new EventListener() {
				
				@Override
				public void onEvent(Event event) throws Exception {
					Map data = (Map)event.getData();
					Chart.Type chartType = (Chart.Type) insertChartAtDialog.getAttribute("chartType");
					insertChartAtDialog.setAttribute("chartType", null); //clear it
					if (data != null && chartType != null) {
						Integer col = (Integer) data.get("column");
						Integer row = (Integer) data.get("row");
						getWorkbookCtrl().addChart(row, col, chartType);
					}
				}
			});
		}
		insertChartAtDialog.setAttribute("chartType", chartType); //put into dialog attribute for onClose() listener
		insertChartAtDialog.fireOnOpen(null);
	}
}
