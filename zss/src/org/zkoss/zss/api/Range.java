package org.zkoss.zss.api;

import java.util.List;

import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Color;
import org.zkoss.zss.api.model.Font;
import org.zkoss.zss.api.model.CellStyle.BorderType;
import org.zkoss.zss.api.model.Chart;
import org.zkoss.zss.api.model.Chart.Grouping;
import org.zkoss.zss.api.model.Chart.LegendPosition;
import org.zkoss.zss.api.model.Chart.Type;
import org.zkoss.zss.api.model.ChartData;
import org.zkoss.zss.api.model.Font.Boldweight;
import org.zkoss.zss.api.model.Font.TypeOffset;
import org.zkoss.zss.api.model.Font.Underline;
import org.zkoss.zss.api.model.Hyperlink.HyperlinkType;
import org.zkoss.zss.api.model.Picture;
import org.zkoss.zss.api.model.Picture.Format;
import org.zkoss.zss.api.model.Sheet;

/**
 * 1.Range is not handling the protection issue, if you have handle it yourself before calling the api(by calling {@code #isProtected()})
 * @author dennis
 * @since 3.0.0
 */
public interface Range {
	
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
	
	public Book getBook();
	
	public Sheet getSheet();
	
	public int getColumn();
	public int getRow();
	public int getLastColumn();
	public int getLastRow();
	
	public StyleHelper getStyleHelper();

	public boolean isProtected();
	
	public void clearContents();

	public void clearStyles();

	/**
	 * @param dest the destination 
	 * @return true if paste successfully, past to a protected sheet with any
	 *         locked cell in the destination range will always cause past fail.
	 */
	public boolean paste(Range dest);
	
	/**
	 * @param dest the destination 
	 * @param transpose TODO
	 * @return true if paste successfully, past to a protected sheet with any
	 *         locked cell in the destination range will always cause past fail.
	 */
	public boolean pasteSpecial(Range dest,PasteType type,PasteOperation op,boolean skipBlanks,boolean transpose);


	public void setStyle(CellStyle nstyle);
	
	public void sync(RangeRunner run);
	/**
	 * visit all cells in this range, make sure you call this in a limited range, 
	 * don't use it for all row/column selection, it will spend much time to iterate the cell 
	 * @param visitor the visitor 
	 * @param create create cell if it doesn't exist, if it is true, it will also lock the sheet
	 */
	public void visit(final CellVisitor visitor);
	
	public void applyBorders(ApplyBorderType type,BorderType borderType,String htmlColor);

	public boolean hasMergeCell();
	
	public void merge(boolean across);
	
	public void unMerge();

	public Range getShiftedRange(int rowOffset,int colOffset);
	
	public Range getCellRange(int rowOffset,int colOffset);
	
	/**
	 *  Return a range that represents all columns and between the first-row and last-row of this range
	 **/
	public Range getRowRange();
	
	/**
	 *  Return a range that represents all rows and between the first-column and last-column of this range
	 **/
	public Range getColumnRange();
	
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
	
	public void sort(Range index1,
			boolean desc1,
			boolean header, 
			/*int orderCustom, //not implement*/
			boolean matchCase, 
			boolean sortByRows, 
			/*int sortMethod, //not implement*/
			SortDataOption dataOption1,
			Range index2,boolean desc2,SortDataOption dataOption2,
			Range index3,boolean desc3,SortDataOption dataOption3);
	
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
	
	public void fill(Range dest,AutoFillType fillType);
	
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
	public CellStyle getCellStyle();
	
	public Picture addPicture(SheetAnchor anchor,byte[] image,Format format);
	
	public List<Picture> getPictures();
	
	public void deletePicture(Picture picture);
	
	public void movePicture(SheetAnchor anchor,Picture picture);
	
	//currently, we only support to modify chart in XSSF
	public Chart addChart(SheetAnchor anchor,ChartData data,Type type, Grouping grouping, LegendPosition pos);
	
	public List<Chart> getCharts();
	
	//currently, we only support to modify chart in XSSF
	public void deleteChart(Chart chart);
	
	//currently, we only support to modify chart in XSSF
	public void moveChart(SheetAnchor anchor,Chart chart);
	
	public Sheet createSheet(String name);
	
	public void deleteSheet();
	
	
	/**
	 * a cell style helper to create style relative object for cell
	 * @author dennis
	 */
	public interface StyleHelper {

		/**
		 * create a new cell style and clone attribute from src if it is not null
		 * @param src the source to clone, could be null
		 * @return the new cell style
		 */
		public CellStyle createCellStyle(CellStyle src);

		/**
		 * create a new font and clone attribute from src if it is not null
		 * @param src the source to clone, could be null
		 * @return the new font
		 */
		public Font createFont(Font src);
		
		/**
		 * create a color object from a htmlColor expression
		 * @param htmlColor html color expression, ex. #FF00FF
		 * @return a Color object
		 */
		public Color createColorFromHtmlColor(String htmlColor);
		
		/**
		 * find the font with given condition
		 * @param boldweight
		 * @param color
		 * @param fontHeight
		 * @param fontName
		 * @param italic
		 * @param strikeout
		 * @param typeOffset
		 * @param underline
		 * @return null if not found
		 */
		public Font findFont(Boldweight boldweight, Color color,
				short fontHeight, String fontName, boolean italic,
				boolean strikeout, TypeOffset typeOffset, Underline underline);
	}

}
