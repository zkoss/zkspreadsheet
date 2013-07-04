package org.zkoss.zss.jsfessentials;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.Date;

import javax.faces.application.FacesMessage;
import javax.faces.bean.ManagedBean;
import javax.faces.bean.RequestScoped;
import javax.faces.context.FacesContext;
import javax.faces.context.PartialResponseWriter;

import org.zkoss.json.JSONObject;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.jsf.Action;
import org.zkoss.zss.jsf.Update;
import org.zkoss.zss.jsf.ui.JSFUpdate;

@ManagedBean
@RequestScoped
public class ApplicationForLeave {

	Book book;

	public ApplicationForLeave() {
	}

	String dateFormat = "yyyy/MM/dd";

	public Book getBook() {
		if (book != null) {
			return book;
		}
		try {
			URL bookUrl = FacesContext.getCurrentInstance()
					.getExternalContext()
					.getResource("/WEB-INF/books/application_for_leave.xlsx");
			book = Importers.getImporter().imports(bookUrl, "app4leave");
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}

		Sheet sheet = book.getSheetAt(0);

		// reset sample data
		// you can use a cell reference to get a range
		Range from = Ranges.range(sheet, "E5");// Ranges.range(sheet,"From");
		// or you can use a name to get a range (the named range has to be set in book);
		Range to = Ranges.rangeByName(sheet, "To");
		Range reason = Ranges.rangeByName(sheet, "Reason");
		Range applicant = Ranges.rangeByName(sheet, "Applicant");
		Range requestDate = Ranges.rangeByName(sheet, "RequestDate");

		// use range api to set the cell data
		from.setCellEditText(DateUtil.tomorrow(0, dateFormat));
		to.setCellEditText(DateUtil.tomorrow(0, dateFormat));
		reason.setCellEditText("");
		applicant.setCellEditText("");
		requestDate.setCellEditText(DateUtil.today(dateFormat));

		return book;
	}

	public void setBook(Book book) {
		this.book = book;
	}

	Update update;

	public void setUpdate(Update zkupdate) {
		this.update = zkupdate;
	}
	
	public Update getUpdate(){
		return update;
	}

	public void doReset() {
		
		//use zkupdate to execute the action inside zk context
		//so the sparedsheet can get the update of book automatically
		update.execute(new Action() {
			public void execute() {
				Sheet sheet = book.getSheetAt(0);

				// reset sample data
				// you can use a cell reference to get a range
				Range from = Ranges.range(sheet, "E5");// Ranges.range(sheet,"From");
				// or you can use a name to get a range (the named rnage has to be
				// set in book);
				Range to = Ranges.rangeByName(sheet, "To");
				Range reason = Ranges.rangeByName(sheet, "Reason");
				Range applicant = Ranges.rangeByName(sheet, "Applicant");
				Range requestDate = Ranges.rangeByName(sheet, "RequestDate");

				// use range api to set the cell data
				from.setCellEditText(DateUtil.tomorrow(0, dateFormat));
				to.setCellEditText(DateUtil.tomorrow(0, dateFormat));
				reason.setCellEditText("");
				applicant.setCellEditText("");
				requestDate.setCellEditText(DateUtil.today(dateFormat));
			}
		});
		addMessage("Reset book");
		
//		PartialResponseWriter responseWriter = FacesContext.getCurrentInstance().getPartialViewContext().getPartialResponseWriter();
//        try {
//			responseWriter.startDocument();
//		
//	        responseWriter.startEval();
//	        responseWriter.write("alert('abc')");
//	        responseWriter.endEval();
//	        responseWriter.endDocument();
//	        responseWriter.flush();
//	        responseWriter.close();
//        } catch (IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
	}
	
	private void addMessage(String message){
		FacesContext.getCurrentInstance().addMessage(null, new FacesMessage(message));
	}

	public void doOk() {
		//access cell data
		update.execute(new Action() {
			public void execute() {
				Sheet sheet = book.getSheetAt(0);

				Date from = Ranges.rangeByName(sheet,"From").getCellData().getDateValue();
				Date to = Ranges.rangeByName(sheet,"To").getCellData().getDateValue();
				String reason = Ranges.rangeByName(sheet,"Reason").getCellData().getStringValue();
				Double total = Ranges.rangeByName(sheet,"Total").getCellData().getDoubleValue();
				String applicant = Ranges.rangeByName(sheet,"Applicant").getCellData().getStringValue();
				Date requestDate = Ranges.rangeByName(sheet,"RequestDate").getCellData().getDateValue();
				
				if(from == null){
					addMessage("FROM is empty");
				}else if(to == null){
					addMessage("TO is empty");
				}else if(total==null || total.intValue()<0){
					addMessage("TOTAL small than 1");
				}else if(reason == null){
					addMessage("REASON is empty");
				}else if(applicant == null){
					addMessage("APPLICANT is empty");
				}else if(requestDate == null){
					addMessage("REQUEST DATE is empty");
				}else{
					//Handle your business logic here 
					
					addMessage("Your request are sent, following is your data");

					addMessage("From :" +from);//can't pass as data, use long for time
					addMessage("To :" + to);//can't pass as data, use long for time
					addMessage("Reason :"+ reason);
					addMessage("Total :"+ total.intValue());//we just need int
					addMessage("Applicant :"+ applicant);
					addMessage("RequestDate :"+ requestDate.getTime());
					
					//You can also store the book, and load it back later by export it to a file
					Exporter exporter = Exporters.getExporter();
					FileOutputStream fos = null;
					try {
						File temp = File.createTempFile("app4leave_", ".xlsx");
						fos = new FileOutputStream(temp); 
						exporter.export(sheet.getBook(), fos);
						System.out.println("file save at "+temp.getAbsolutePath());
						
						addMessage("Archive "+ temp.getName());
					} catch (IOException e) {
						e.printStackTrace();
					} finally{
						if(fos!=null)
							try {
								fos.close();
							} catch (IOException e) {//eat
							}
					}
				}
			}
		});
		
				
	}
}
