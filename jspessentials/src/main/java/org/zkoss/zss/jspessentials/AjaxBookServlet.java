package org.zkoss.zss.jspessentials;

import java.io.IOException;
import java.io.Writer;
import java.security.Principal;
import java.util.Date;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.zkoss.json.JSONObject;
import org.zkoss.zk.ui.Desktop;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zssex.ui.AjaxUpdateBridge;

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
		//action parameter from ajax request, you have to pass it in ajax request
		final String dtid = req.getParameter("dtid");
		final String ssuuid = req.getParameter("ssuuid");
		
		final String action = req.getParameter("action");
		
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
						//update cell data
						
						//reset sample data
						Ranges.range(sheet,"D4").setCellEditText(DateUtil.tomorrow(0,"yyyy/MM/dd"));//from
						Ranges.range(sheet,"D5").setCellEditText(DateUtil.tomorrow(0,"yyyy/MM/dd"));//to
						Ranges.range(sheet,"C6").setCellEditText("");//reason
						
						Principal p = _request.getUserPrincipal();
						if(p!=null){
							Ranges.range(sheet,"C7").setCellEditText(p.getName());//applicant name for base servlet auth
						}else{
							Ranges.range(sheet,"C7").setCellEditText("");//applicant name for base servlet auth
						}
						
					}else if("check".equals(action)){
						//access cell data
						Date from = Ranges.range(sheet,"D4").getCellData().getDateValue();
						Date to = Ranges.range(sheet,"D5").getCellData().getDateValue();
						String reason = Ranges.range(sheet,"C6").getCellData().getStringValue();
						Double total = Ranges.range(sheet,"F4").getCellData().getDoubleValue();
						String applicant = Ranges.range(sheet,"C7").getCellData().getStringValue();
						
						Range r = Ranges.range(sheet,"D4");
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
						}else{
							result.put("validation", true);
						}
						
					}else{
						result.put("error", "unknow action");
					}
				}else{
					result.put("error", "can't find any zk sparedsheet with uuid "+ssuuid);
				}
			}
		};
		
		if(sb.hasDesktop()){
			sb.process(result);
		}else{
			result.put("error", "can't find any zk desktop with dtid "+dtid);
		}
		
		Writer w = resp.getWriter();
		w.append(result.toJSONString());	
	}
}
