package org.zkoss.zss.api.impl;

import java.io.File;
import java.io.IOException;

import org.zkoss.image.AImage;
import org.zkoss.zss.api.AreaRef;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.SheetOperationUtil;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Chart;
import org.zkoss.zss.api.model.ChartData;
import org.zkoss.zss.api.model.Picture;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.Chart.Grouping;
import org.zkoss.zss.api.model.Chart.LegendPosition;
import org.zkoss.zssex.api.ChartDataUtil;

/**
 * all method implementation to test chart & picture operation
 * @author kuro
 */
public class ChartPictureTestBase {
	
	protected void testDeletePicture(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Picture picture = SheetOperationUtil.addPicture(Ranges.range(sheet), new AImage(new File(ChartPictureTestBase.class.getResource("").getPath() + "book/zklogo.png")));
		SheetOperationUtil.deletePicture(Ranges.range(sheet), picture);
	}
	
	protected void testMovePicture(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Picture picture = SheetOperationUtil.addPicture(Ranges.range(sheet), new AImage(new File(ChartPictureTestBase.class.getResource("").getPath() + "book/zklogo.png")));
		SheetOperationUtil.movePicture(Ranges.range(sheet), picture, 10, 30);
	}
	
	protected void testAddPicture(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		SheetOperationUtil.addPicture(Ranges.range(sheet), new AImage(new File(ChartPictureTestBase.class.getResource("").getPath() + "book/zklogo.png")));
	}
	
	protected void testMoveChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		ChartData cd1 = ChartDataUtil.getChartData(sheet, new AreaRef(4,1,14,1), Chart.Type.LINE);
		Chart chart = SheetOperationUtil.addChart(Ranges.range(sheet, "A1"), cd1, Chart.Type.LINE, Grouping.STANDARD, LegendPosition.TOP);
		SheetOperationUtil.moveChart(Ranges.range(sheet), chart, 10, 20);
	}
	
	protected void testDeleteChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		ChartData cd1 = ChartDataUtil.getChartData(sheet, new AreaRef(4,1,14,1), Chart.Type.LINE);
		Chart chart = SheetOperationUtil.addChart(Ranges.range(sheet, "A1"), cd1, Chart.Type.LINE, Grouping.STANDARD, LegendPosition.TOP);
		SheetOperationUtil.deleteChart(Ranges.range(sheet), chart);
	}
	
	protected void testAddLineChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		ChartData cd1 = ChartDataUtil.getChartData(sheet, new AreaRef(4,1,14,1), Chart.Type.LINE);
		SheetOperationUtil.addChart(Ranges.range(sheet, "A1"), cd1, Chart.Type.LINE, Grouping.STANDARD, LegendPosition.TOP);
	}
	
	protected void testAddBarChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		NChartData cd3 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.BAR);
		SheetOperationUtil.addChart(Ranges.range(sheet, "Q1"), cd3, Chart.Type.BAR, Grouping.STANDARD, LegendPosition.TOP);
	}
	
	protected void testAddAreaChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		ChartData cd4 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.AREA);
		SheetOperationUtil.addChart(Ranges.range(sheet, "Q10"), cd4, Chart.Type.AREA, Grouping.STANDARD, LegendPosition.TOP);
	}
	
	// unsupported 3.0.0 RC
	protected void testAddBubbleChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		ChartData cd5 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.BUBBLE);
		SheetOperationUtil.addChart(Ranges.range(sheet, "Q20"), cd5, Chart.Type.BUBBLE, Grouping.STANDARD, LegendPosition.TOP);
	}
	
	protected void testAddColumnChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		ChartData cd6 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.COLUMN);
		SheetOperationUtil.addChart(Ranges.range(sheet, "Q30"), cd6, Chart.Type.COLUMN, Grouping.STANDARD, LegendPosition.TOP);
	}
	
	protected void testAddPieChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		ChartData cd7 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.PIE);
		SheetOperationUtil.addChart(Ranges.range(sheet, "Q40"), cd7, Chart.Type.PIE, Grouping.STANDARD, LegendPosition.TOP);
	}
	
	// Not Implement 3.0.0 RC
	protected void testAddRadarChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		ChartData cd8 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.RADAR);
		SheetOperationUtil.addChart(Ranges.range(sheet, "Q50"), cd8, Chart.Type.RADAR, Grouping.STANDARD, LegendPosition.TOP);
	}
	
	protected void testAddScatterChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		ChartData cd9 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.SCATTER);
		SheetOperationUtil.addChart(Ranges.range(sheet, "Q60"), cd9, Chart.Type.SCATTER, Grouping.STANDARD, LegendPosition.TOP);
	}
	
	// unsupported 3.0.0 RC
	protected void testAddStockChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		ChartData cd10 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.STOCK);
		SheetOperationUtil.addChart(Ranges.range(sheet, "Q70"), cd10, Chart.Type.STOCK, Grouping.STANDARD, LegendPosition.TOP);
	}
	
	// unsupported 3.0.0 RC
	protected void testAddSurfaceChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		ChartData cd11 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.SURFACE);
		SheetOperationUtil.addChart(Ranges.range(sheet, "Q80"), cd11, Chart.Type.SURFACE, Grouping.STANDARD, LegendPosition.TOP);
	}
	
	protected void testAddDoughnutChart(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("chart-image");
		ChartData cd11 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.DOUGHNUT);
		SheetOperationUtil.addChart(Ranges.range(sheet, "Q90"), cd11, Chart.Type.DOUGHNUT, Grouping.STANDARD, LegendPosition.TOP);
	}
}
