package zss.testapp;


/**
 * A simple data bean.
 * @author Hawk
 *
 */
public class AssetsBean {
    private double liquidAssets = 146504221;
    private double fundInvestment = 23181709;
    private double fixedAssets = 7168392;
    private double intangibleAsset = 221426; 
    private double otherAssets = 2270018;
    
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
}