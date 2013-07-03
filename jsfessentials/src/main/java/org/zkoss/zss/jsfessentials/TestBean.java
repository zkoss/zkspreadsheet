package org.zkoss.zss.jsfessentials;

import java.net.URL;

import javax.faces.bean.ManagedBean;
import javax.faces.bean.RequestScoped;
import javax.faces.bean.SessionScoped;
import javax.faces.context.FacesContext;

import org.zkoss.zkjsf.ui.Action;
import org.zkoss.zkjsf.ui.Update;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

@ManagedBean
@RequestScoped
public class TestBean {

	String message = "Click following test button";

	public String getMessage() {
		return message;
	}

	public void setMessage(String message) {
		this.message = message;
	}
	
	Update update;
	
	public Update getUpdate() {
		return update;
	}

	public void setUpdate(Update update) {
		this.update = update;
	}
	
	Book book;

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

		Range reason = Ranges.rangeByName(sheet, "Reason");

		return book;
	}

	public void setBook(Book book) {
		this.book = book;
	}


	public void doPostBack(){
		this.message = "doPostBack";
		updateBookMessage(this.message);
	}
	
	public void doAjaxNoZss(){
		this.message = "doAjaxNoZss";
		updateBookMessage(this.message);
	}
	
	public void doAjaxZss(){
		this.message = "doAjaxZss";
		updateBookMessage(this.message);
	}
	
	public void doAjaxZssParent(){
		this.message = "doAjaxZssParent";
		updateBookMessage(this.message);
	}

	private void updateBookMessage(final String message){
		update.execute(new Action() {
			public void execute() {
				Sheet sheet = book.getSheetAt(0);
				Range reason = Ranges.rangeByName(sheet, "Reason");
				reason.setCellEditText(message);
			}
		});
	}
	
}
