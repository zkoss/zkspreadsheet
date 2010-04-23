package org.zkoss.zss.app;

import java.io.InputStream;
import java.sql.Connection;
import java.sql.Date;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Path;
import org.zkoss.zk.ui.Session;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.event.SelectEvent;
import org.zkoss.zk.ui.ext.AfterCompose;
import org.zkoss.zss.model.Book;
//import org.zkoss.zss.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Button;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Listcell;
import org.zkoss.zul.Listitem;
import org.zkoss.zul.Window;



/**
 * @author kinda lu
 */
public class ViewWindow extends Window implements AfterCompose {
	Spreadsheet spreadsheet;
	Book book;
	

	public void afterCompose() {
		spreadsheet = new Spreadsheet();
		windowPopupOnceByName("fileOpenViewWin");
		windowPopupOnceByName("menuPrintWin");
        try {
			Class.forName("com.mysql.jdbc.Driver").newInstance();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	private void windowPopupOnceByName(String s){
		Window win = (Window) getFellow(s);
		win.doPopup();
		win.setVisible(false);
	}
	
	public void onOpenFile(ForwardEvent event){
		Listbox flo_files = (Listbox) Path.getComponent("//p2/mainWin/fileOpenViewWin/flo_files");
		
		String filename = flo_files.getSelectedItem().getLabel();
		openFileInSS(filename);

		Window fileOpenWin = (Window) Path.getComponent("//p2/mainWin/fileOpenViewWin");
		fileOpenWin.setVisible(false);
		
		book=spreadsheet.getBook();
		spreadsheet.setMaxrows(200);
		spreadsheet.setMaxcolumns(30);
		Sheet sheet =(Sheet) book.getSheetAt(0);
		spreadsheet.setSelectedSheet(sheet.getSheetName());
		onPrint();
	}
	
	public void openFileInSS(String filename) {
        try
        {
        	Connection con;
        	con=getDBConnection();
        	Statement stmt = con.createStatement();
        	String operation="SELECT spreadsheet FROM `revisionhistory` WHERE stack_level=0 AND filename='"+filename+"' LIMIT 0 , 30";

        	ResultSet targetRs = stmt.executeQuery(operation);

        	if(targetRs.next()){	
        		InputStream iStream=targetRs.getBinaryStream("spreadsheet");

        		openSpreadsheetFromStream(iStream, filename);

        	}
        	targetRs.close();
        	stmt.close();

        }
        catch (SQLException ex)
        {
        	System.out.println("exception: "+ex.getMessage());
        	System.out.println("SQLState: "+ex.getSQLState());
        	System.out.println("errorCode: "+ex.getErrorCode());
        } 
	}
	
	public void onPrint(){

		String printKey=""+System.currentTimeMillis();
		Session session=this.getDesktop().getSession();
		session.setAttribute("zssFromHi"+printKey,spreadsheet);

		Window win = (Window) getFellow("menuPrintWin");
		Button printBtn=(Button) win.getFellow("printBtn");
		printBtn.setTarget("_blank");
		printBtn.setHref("print.zul?printKey="+printKey);

		win.setPosition("parent");
		win.setTop("30px");
		win.setLeft("5px");
		win.doPopup();
	}
	
	public void onOpenFileMenu(ForwardEvent event){
		Window win = (Window) Path.getComponent("//p2/mainWin/fileOpenViewWin");

		reloadMenu("open");
		try {
			win.setTop("30px");
			win.setLeft("5px");
			win.doPopup();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void reloadMenu(String target){
		Connection con;

		try
        {
            con=DriverManager.getConnection("jdbc:mysql://localhost:3306/zss","root","rootzk");
            Statement stmt = con.createStatement();
            String operation;
            operation="SELECT time,who,comment,filename FROM `revisionhistory` WHERE stack_level=0 ORDER BY stack_level DESC,filename LIMIT 0 , 30";
           // p("...in MainWindow: "+operation);
            ResultSet rs = stmt.executeQuery(operation);
            Listbox listbox=null;
            if(target.equals("open"))
            	listbox = (Listbox)Path.getComponent("//p2/mainWin/fileOpenViewWin/flo_files");
            if(target.equals("delete"))
            	listbox = (Listbox)Path.getComponent("//p2/mainWin/fileDeleteWin/fld_files");
            if(target.equals("export"))
            	listbox = (Listbox)Path.getComponent("//p2/mainWin/fileExportWin/fle_files");
			
			List childList = listbox.getChildren();
			//remain the list head component
			while(childList.size()>1)
				listbox.removeChild((Component)childList.get(1));
			
            while(rs.next()){
            	//System.out.println("in rs.next");
            	Listitem newListItem = new Listitem(); 
            	Date date = rs.getDate("time");
            	SimpleDateFormat sdf=new SimpleDateFormat("M/dd h");
            	newListItem.appendChild(new Listcell(rs.getString("filename")));
            	newListItem.appendChild(new Listcell(rs.getString("who")));
            	newListItem.appendChild(new Listcell(String.valueOf(sdf.format(date))));
            	listbox.appendChild(newListItem);
            }
            //close the jdbc connection
            stmt.close();
            con.close();
            
            listbox.setSelectedIndex(0);
        }
        catch (SQLException ex)
        {
            System.out.println("exception: "+ex.getMessage());
            System.out.println("SQLState: "+ex.getSQLState());
            System.out.println("errorCode: "+ex.getErrorCode());
        }
	}
	
	private Connection getDBConnection(){
		try {
			return DriverManager.getConnection("jdbc:mysql://localhost:3306/zss","root","rootzk");
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	private void p(String s){
		System.out.println(s);
	}
	
	public void openSpreadsheetFromStream(InputStream iStream, String src){
		spreadsheet.setBookFromStream(iStream, src);
	}
}








