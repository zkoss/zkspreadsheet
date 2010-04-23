package org.zkoss.zss.demo;

import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.zkoss.zss.demo.ComputerBean;

import org.zkoss.zk.ui.Executions;

import au.com.bytecode.opencsv.CSVReader;

public class ReportProvider {
	public static QuarterBean queryQuarterBean(int quarter, QuarterBean dataBean) {
		if(quarter == 0) {
			dataBean.setItem("Quarter");
		} else {
			dataBean.setItem("Quarter " + quarter);
		}
		if(quarter == 1) {
			dataBean.setLiquidAssets(new Double(146504221));
			dataBean.setFundInvestment(new Double(23181709));
			dataBean.setFixedAssets(new Double(7168392));
			dataBean.setIntangibleAsset(new Double(221426));
			dataBean.setOtherAssets(new Double(2270018));
			dataBean.setCurrentLiabilities(new Double(102515784));
			dataBean.setLongTermLiabilities(new Double(3000));
			dataBean.setOtherLiabilities(new Double(456175));
			dataBean.setCapitalStock(new Double(33630080));
			dataBean.setCapitalSurplus(new Double(7127901));
			dataBean.setRetainedEarnings(new Double(34420905));
			dataBean.setOtherEquity(new Double(1826731));
			dataBean.setTreasuryStock(new Double(-634810));
		} else if(quarter == 2) {
			dataBean.setLiquidAssets(new Double(160793618));
			dataBean.setFundInvestment(new Double(26019684));
			dataBean.setFixedAssets(new Double(7106496));
			dataBean.setIntangibleAsset(new Double(176485));
			dataBean.setOtherAssets(new Double(2183907));
			dataBean.setCurrentLiabilities(new Double(121447738));
			dataBean.setLongTermLiabilities(new Double(3000));
			dataBean.setOtherLiabilities(new Double(722960));
			dataBean.setCapitalStock(new Double(34752682));
			dataBean.setCapitalSurplus(new Double(7133397));
			dataBean.setRetainedEarnings(new Double(29056591));
			dataBean.setOtherEquity(new Double(3811127));
			dataBean.setTreasuryStock(new Double(-647305));
		} else if(quarter == 3) {
			dataBean.setLiquidAssets(new Double(187967615));
			dataBean.setFundInvestment(new Double(27159093));
			dataBean.setFixedAssets(new Double(7047513));
			dataBean.setIntangibleAsset(new Double(182348));
			dataBean.setOtherAssets(new Double(2166805));
			dataBean.setCurrentLiabilities(new Double(142212438));
			dataBean.setLongTermLiabilities(new Double(3000));
			dataBean.setOtherLiabilities(new Double(717980));
			dataBean.setCapitalStock(new Double(34752682));
			dataBean.setCapitalSurplus(new Double(7234727));
			dataBean.setRetainedEarnings(new Double(34300237));
			dataBean.setOtherEquity(new Double(4667500));
			dataBean.setTreasuryStock(new Double(634810));
		} else {
			dataBean.setLiquidAssets(new Double(0));
			dataBean.setFundInvestment(new Double(0));
			dataBean.setFixedAssets(new Double(0));
			dataBean.setIntangibleAsset(new Double(0));
			dataBean.setOtherAssets(new Double(0));
			dataBean.setCurrentLiabilities(new Double(0));
			dataBean.setLongTermLiabilities(new Double(0));
			dataBean.setOtherLiabilities(new Double(0));
			dataBean.setCapitalStock(new Double(0));
			dataBean.setCapitalSurplus(new Double(0));
			dataBean.setRetainedEarnings(new Double(0));
			dataBean.setOtherEquity(new Double(0));
			dataBean.setTreasuryStock(new Double(0));
		}
		return dataBean;
	}

	public static List queryComputerBean() throws IOException {
		File csv = new File(Executions.getCurrent().getDesktop().getWebApp().getRealPath("/WEB-INF/xls/demo/data.csv"));
		CSVReader reader = new CSVReader(new FileReader(csv));
		List data = new ArrayList();
		String[] nextLine;
	    while ((nextLine = reader.readNext()) != null) {
	    	ComputerBean computerBean = new ComputerBean();
	    	computerBean.setId(nextLine[0]);
	    	computerBean.setProduct(nextLine[1]);
	    	computerBean.setBrand(nextLine[2]);
	    	computerBean.setModel(nextLine[3]);
	    	computerBean.setSerialNumber(nextLine[4]);
	    	computerBean.setDate(nextLine[5]);
	    	computerBean.setWarrantyTime(nextLine[6]);
	    	computerBean.setCost(Double.valueOf(nextLine[7]).doubleValue());
	    	computerBean.setOs(nextLine[8]);
	    	computerBean.setSalvage(Double.valueOf(nextLine[9]).doubleValue());
	        data.add(computerBean);
	    }
	    return data;
	}
}
