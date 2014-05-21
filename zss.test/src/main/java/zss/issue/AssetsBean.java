package zss.issue;

import java.util.ArrayList;
import java.util.List;

//import org.springframework.context.annotation.Scope;
//import org.springframework.context.annotation.ScopedProxyMode;
//import org.springframework.stereotype.Component;

/**
 * A simple data bean.
 * @author Hawk
 *
 */
//@Component
//@Scope(value="session",  proxyMode=ScopedProxyMode.TARGET_CLASS)
public class AssetsBean {
    private double liquidAssets = 146504221;
    private double fundInvestment = 23181709;
    private double fixedAssets = 7168392;
    private double intangibleAsset = 221426; 
    private double otherAssets = 2270018;
    private List<Double> list = new ArrayList<Double>();
    
	public double getLiquidAssets() {
		return liquidAssets;
	}
	public void setLiquidAssets(double liquidAssets) {
		this.liquidAssets = liquidAssets;
	}
	public double getFundInvestment() {
		return fundInvestment;
	}
	public void setFundInvestment(double fundInvestment) {
		this.fundInvestment = fundInvestment;
	}
	public double getFixedAssets() {
		return fixedAssets;
	}
	public void setFixedAssets(double fixedAssets) {
		this.fixedAssets = fixedAssets;
	}
	public double getIntangibleAsset() {
		return intangibleAsset;
	}
	public void setIntangibleAsset(double intangibleAsset) {
		this.intangibleAsset = intangibleAsset;
	}
	public double getOtherAssets() {
		return otherAssets;
	}
	public void setOtherAssets(double otherAssets) {
		this.otherAssets = otherAssets;
	}
	public List<Double> getList() {
		if (list.isEmpty()) {
			list.add(new Double(12111));
			list.add(new Double(12113));
			list.add(new Double(12114));
			list.add(new Double(12116));
		}
		return list;
	}
	public void setList(List<Double> list) {
		this.list = list;
	}
}