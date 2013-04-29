package org.zkoss.zss.api;

import java.util.List;

import org.zkoss.zss.api.model.NBook;
import org.zkoss.zss.api.model.NCellStyle;
import org.zkoss.zss.api.model.NCellStyle.BorderType;
import org.zkoss.zss.api.model.NChart;
import org.zkoss.zss.api.model.NChart.Grouping;
import org.zkoss.zss.api.model.NChart.LegendPosition;
import org.zkoss.zss.api.model.NChart.Type;
import org.zkoss.zss.api.model.NChartData;
import org.zkoss.zss.api.model.NHyperlink.HyperlinkType;
import org.zkoss.zss.api.model.NPicture;
import org.zkoss.zss.api.model.NPicture.Format;
import org.zkoss.zss.api.model.NSheet;

/**
 * 1.Range is not handling the protection issue, if you have handle it yourself before calling the api(by calling {@code #isProtected()})
 * @author dennis
 *
 */
public interface NRange {
	
	enum SyncLevel{
		BOOK,
		NONE//for you just visit and do nothing
	}
	
	public enum PasteType{
		PASTE_ALL,
		PASTE_ALL_EXCEPT_BORDERS,
		PASTE_COLUMN_WIDTHS,
		PASTE_COMMENTS,
		PASTE_FORMATS/*all formats*/,
		PASTE_FORMULAS/*include values and formulas*/,
		PASTE_FORMULAS_AND_NUMBER_FORMATS,
		PASTE_VALIDATAION,
		PASTE_VALUES,
		PASTE_VALUES_AND_NUMBER_FORMATS;
	}
	
	public enum PasteOperation{
		PASTEOP_ADD,
		PASTEOP_SUB,
		PASTEOP_MUL,
		PASTEOP_DIV,
		PASTEOP_NONE;
	}
	
	public enum ApplyBorderType{
		FULL,
		EDGE_BOTTOM,
		EDGE_RIGHT,
		EDGE_TOP,
		EDGE_LEFT,
		OUTLINE,
		INSIDE,
		INSIDE_HORIZONTAL,
		INSIDE_VERTICAL,
		DIAGONAL,
		DIAGONAL_DOWN,
		DIAGONAL_UP
	}
	
	/** Shift direction of insert and delete**/
	public enum InsertShift{
		DEFAULT,
		RIGHT,
		DOWN
	}
	/** copy origin of insert and delete**/
	public enum InsertCopyOrigin{
		NONE,
		LEFT_ABOVE,
		RIGHT_BELOW,
	}
	/** Shift direction of insert and delete**/
	public enum DeleteShift{
		DEFAULT,
		LEFT,
		UP
	}
	
	public enum SortDataOption{
		TEXT_AS_NUMBERS
	}
	
	public enum AutoFilterOperation{
		AND,
		BOTTOM10,
		BOTOOM10PERCENT,
		OR,
		TOP10,
		TOP10PERCENT,
		VALUES
	}
	
	public enum AutoFillType{
		COPY,
		DAYS,
		DEFAULT,
		FORMATS,
		MONTHS,
		SERIES,
		VALUES,
		WEEKDAYS,
		YEARS,
		GROWTH_TREND,
		LINER_TREND
	}


	public void setSyncLevel(SyncLevel syncLevel);
	
	public NBook getBook();
	
	public NSheet getSheet();
	
	public int getColumn();
	public int getRow();
	public int getLastColumn();
	public int getLastRow();
	
	public NCellStyleHelper getCellStyleHelper();

	public boolean isProtected();
	
	public void clearContents();

	public void clearStyles();

	/**
	 * @param dest the destination 
	 * @return true if paste successfully, past to a protected sheet with any
	 *         locked cell in the destination range will always cause past fail.
	 */
	public boolean paste(NRange dest);
	
	/**
	 * @param dest the destination 
	 * @param transpose TODO
	 * @return true if paste successfully, past to a protected sheet with any
	 *         locked cell in the destination range will always cause past fail.
	 */
	public boolean pasteSpecial(NRange dest,PasteType type,PasteOperation op,boolean skipBlanks,boolean transpose);


	public void setStyle(NCellStyle nstyle);
	
	public void sync(NRangeRunner run);
	/**
	 * visit all cells in this range, make sure you call this in a limited range, 
	 * don't use it for all row/column selection, it will spend much time to iterate the cell 
	 * @param visitor the visitor 
	 * @param create create cell if it doesn't exist, if it is true, it will also lock the sheet
	 */
	public void visit(final NCellVisitor visitor);
	
	public void applyBorders(ApplyBorderType type,BorderType borderType,String htmlColor);

	public boolean hasMergeCell();
	
	public void merge(boolean across);
	
	public void unMerge();

	public NRange getShiftedRange(int rowOffset,int colOffset);
	
	public NRange getCellRange(int rowOffset,int colOffset);
	
	/**
	 *  Return a range that represents all columns and between the first-row and last-row of this range
	 **/
	public NRange getRowRange();
	
	/**
	 *  Return a range that represents all rows and between the first-column and last-column of this range
	 **/
	public NRange getColumnRange();
	
	/**
	 * Check if this range represents a whole column, which mean all rows are included, 
	 */
	public boolean isWholeColumn();
	/**
	 * Check if this range represents a whole row, which mean all column are included, 
	 */
	public boolean isWholeRow();
	/**
	 * Check if this range represents a whole sheet, which mean all column and row are included, 
	 */
	public boolean isWholeSheet();
	
	public void insert(InsertShift shift,InsertCopyOrigin copyOrigin);
	
	public void delete(DeleteShift shift);
	
	public void sort(boolean desc);
	
	public void sort(boolean desc,
			boolean header, 
			boolean matchCase, 
			boolean sortByRows, 
			SortDataOption dataOption);
	
	public void sort(NRange index1,
			boolean desc1,
			boolean header, 
			/*int orderCustom, //not implement*/
			boolean matchCase, 
			boolean sortByRows, 
			/*int sortMethod, //not implement*/
			SortDataOption dataOption1,
			NRange index2,boolean desc2,SortDataOption dataOption2,
			NRange index3,boolean desc3,SortDataOption dataOption3);
	
	/** check if auto filter is enable or not.**/
	public boolean isAutoFilterEnabled();
	
	/** enable/disable autofilter of the sheet**/
	public void enableAutoFilter(boolean enable);
	/** enable filter with condition **/
	//TODO have to review this after I know more detail
	public void enableAutoFilter(int field, AutoFilterOperation filterOp, Object criteria1, Object criteria2, Boolean visibleDropDown);
	
	/** clear condition of filter, show all the data**/
	public void resetAutoFilter();
	/** apply the filter to filter again**/
	public void applyAutoFilter();
	
	/** enable sheet protection and apply a password**/
	public void protectSheet(String password);
	
	public void fill(NRange dest,AutoFillType fillType);
	
	public void fillDown();
	
	public void fillLeft();
	
	public void fillUp();
	
	public void fillRight();
	
	/** shift this range with a offset row and column**/
	public void shift(int rowOffset,int colOffset);
	
	public String getEditText();
	
	public void setEditText(String editText);
	
	
	//TODO need to verify the object type
	public Object getValue();
	
	public void setDisplaySheetGridlines(boolean enable);
	
	public boolean isDisplaySheetGridlines();
	
	public void setHidden(boolean hidden);
	
	public void setHyperlink(HyperlinkType type,String address,String displayLabel);
	
	public void setSheetName(String name);
	
	public String getSheetName();
	
	public void setSheetOrder(int pos);
	
	public int getSheetOrder();
	
	public void setValue(Object value);
	
	

	/**
	 * get the first cell style of this range
	 * 
	 * @return cell style if cell is exist, the check row style and column cell style if cell not found, if row and column style is not exist, then return default style of sheet
	 */
	public NCellStyle getCellStyle();
	
	public NPicture addPicture(NSheetAnchor anchor,byte[] image,Format format);
	
	public List<NPicture> getPictures();
	
	public void deletePicture(NPicture picture);
	
	public void movePicture(NSheetAnchor anchor,NPicture picture);
	
	//currently, we only support to modify chart in XSSF
	public NChart addChart(NSheetAnchor anchor,NChartData data,Type type, Grouping grouping, LegendPosition pos);
	
	public List<NChart> getCharts();
	
	//currently, we only support to modify chart in XSSF
	public void deleteChart(NChart chart);
	
	//currently, we only support to modify chart in XSSF
	public void moveChart(NSheetAnchor anchor,NChart chart);
	
	public NSheet createSheet(String name);
	
	public void deleteSheet();
}
