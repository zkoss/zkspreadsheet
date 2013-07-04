package org.zkoss.zss.jspessentials;

import java.net.URL;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.jsp.JSPBookProvider;

public class DemoBookProvider implements JSPBookProvider{

	String dateFormat =  "yyyy/MM/dd";
	
	public Book loadBook(ServletContext servletContext, HttpServletRequest request, HttpServletResponse res) {
		
		Book book;
		try {
			URL bookUrl = servletContext.getResource("/WEB-INF/books/application_for_leave.xlsx");
			book = Importers.getImporter().imports(bookUrl, "app4leave");
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
		
		Sheet sheet = book.getSheetAt(0);
		
		//reset sample data
		//you can use a cell reference to get a range
		Range from = Ranges.range(sheet,"E5");//Ranges.range(sheet,"From");
		//or you can use a name to get a range (the named rnage has to be set in book);
		Range to = Ranges.rangeByName(sheet,"To");
		Range reason = Ranges.rangeByName(sheet,"Reason");
		Range applicant = Ranges.rangeByName(sheet,"Applicant");
		Range requestDate = Ranges.rangeByName(sheet,"RequestDate");
		
		//use range api to set the cell data
		from.setCellEditText(DateUtil.tomorrow(0,dateFormat));
		to.setCellEditText(DateUtil.tomorrow(0,dateFormat));
		reason.setCellEditText("");
		applicant.setCellEditText("");
		requestDate.setCellEditText(DateUtil.today(dateFormat));
		
		return book;
	}

}
