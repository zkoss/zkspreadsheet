package org.zkoss.zss.jspessentials;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.Writer;
import java.util.Date;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.zkoss.json.JSONObject;
import org.zkoss.zk.ui.Desktop;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zssex.ui.AjaxUpdateBridge;
/**
 * a servlet to handle ajax request and return the result
 * @author dennis
 *
 */
public class AjaxBookServlet extends HttpServlet{
	private static final long serialVersionUID = 1L;

	@Override
	protected void doGet(HttpServletRequest req, HttpServletResponse resp)
			throws ServletException, IOException {
		doPost(req, resp);
	}

	@Override
	protected void doPost(HttpServletRequest req, HttpServletResponse resp)
			throws ServletException, IOException {
		//set encoding
		req.setCharacterEncoding("UTF-8");
		resp.setCharacterEncoding("UTF-8");
		
		//parameter from ajax request, you have to pass it in ajax request
		final String dtid = req.getParameter("dtid");//necessary parameter to get zk server side desktop
		final String ssuuid = req.getParameter("ssuuid");//necessary parameter to get zk server side spreadsheet
		
		final String action = req.getParameter("action");
		
		//prepare a json result object, it can contain your ajax result and also the necessary zk component update result
		JSONObject result = new JSONObject();
		result.put("action", action);//set back for client to check action result, it depends on your logic.
		
		//use bridge utility abstract class to wrap zk in servlet request and get access and response result
		AjaxUpdateBridge sb = new AjaxUpdateBridge(getServletContext(), req, resp, dtid) {
			@Override
			protected void process(JSONObject result, Desktop desktop) {
				Spreadsheet ss = (Spreadsheet)desktop.getComponentByUuidIfAny(ssuuid);
				if(ss!=null){
					Book book = ss.getBook();
					Sheet sheet = book.getSheetAt(0);
					
					if("reset".equals(action)){
						handleReset(sheet,result);
						
					}else if("check".equals(action)){
						handleCheck(sheet,result);
					}else{
						result.put("error", "unknow action");
					}
				}else{
					result.put("error", "can't find any zk sparedsheet with uuid "+ssuuid);
				}
			}
		};
		
		if(sb.hasDesktop()){
			/**
			 * zk update result will put in given json object, the default field zk puts in is 'zkjs'.
			 * then in your ajax response handler, it should 'eval' this to get correct zk ui update 
			 */
			sb.process(result);
			
		}else{
			result.put("error", "can't find any zk desktop with dtid "+dtid);//your ajax handling logic
		}
		
		Writer w = resp.getWriter();
		w.append(result.toJSONString());	
	}
	

	private void handleReset(Sheet sheet, JSONObject result) {
		final String dateFormat =  "yyyy/MM/dd";
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
	}
	
	private void handleCheck(Sheet sheet, JSONObject result) {
		//access cell data
		Date from = Ranges.rangeByName(sheet,"From").getCellData().getDateValue();
		Date to = Ranges.rangeByName(sheet,"To").getCellData().getDateValue();
		String reason = Ranges.rangeByName(sheet,"Reason").getCellData().getStringValue();
		Double total = Ranges.rangeByName(sheet,"Total").getCellData().getDoubleValue();
		String applicant = Ranges.rangeByName(sheet,"Applicant").getCellData().getStringValue();
		Date requestDate = Ranges.rangeByName(sheet,"RequestDate").getCellData().getDateValue();
		
		if(from == null){
			result.put("message", "FROM is empty");
		}else if(to == null){
			result.put("message", "TO is empty");
		}else if(total==null || total.intValue()<0){
			result.put("message", "TOTAL small than 1");
		}else if(reason == null){
			result.put("message", "REASON is empty");
		}else if(applicant == null){
			result.put("message", "APPLICANT is empty");
		}else if(requestDate == null){
			result.put("message", "REQUEST DATE is empty");
		}else{
			//Option 1
			//You can handle your business logic here and return a final result for user directly
			
			
			
			//Or option 2, return necessary form data, 
			//so clent can process it by transitional submit that can be handled by spring mvc or strcut
			result.put("valid", true);
			JSONObject form = new JSONObject();
			result.put("form", form);

			form.put("from", from.getTime());//can't pass as data, use long for time
			form.put("to", to.getTime());//can't pass as data, use long for time
			form.put("reason", reason);
			form.put("total", total.intValue());//we just need int
			form.put("applicant", applicant);
			form.put("requestDate", requestDate.getTime());
			
			//You can also store the book, and load it back later by export it to a file
			Exporter exporter = Exporters.getExporter();
			FileOutputStream fos = null;
			try {
				File temp = File.createTempFile("app4leave_", ".xlsx");
				fos = new FileOutputStream(temp); 
				exporter.export(sheet.getBook(), fos);
				System.out.println("file save at "+temp.getAbsolutePath());
				form.put("archive", temp.getName());
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
}
