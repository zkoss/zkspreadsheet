package org.zkoss.zss.essential.advanced;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.VariableResolver;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zk.ui.select.annotation.WireVariable;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Doublebox;

/**
 * Reference Spring bean example.
 * 
 * @author Hawk
 * 
 */

@SuppressWarnings("serial")
@VariableResolver(org.zkoss.zkplus.spring.DelegatingVariableResolver.class)
public class RefSpringBeanComposer extends SelectorComposer<Component> {
	
	@Wire
	private Spreadsheet ss;
	@Wire
	private Doublebox liquidBox;
	@Wire
	private Doublebox fundBox;
	@Wire
	private Doublebox fixedBox;
	@Wire
	private Doublebox intangibleBox;
	@Wire
	private Doublebox otherBox;
	
	@WireVariable
	private AssetsBean assetsBean;
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		Ranges.range(ss.getBook().getSheetAt(0)).protectSheet("");
	}

	@Listen("onChange = doublebox")
	public void update() {
		updateAssetsBean();
		//notify spreadsheet about the bean's change
		Ranges.range(ss.getSelectedSheet()).notifyChange(new String[] {"assetsBean"} );
	}

	/**
	 * load user input to the bean.
	 */
	private void updateAssetsBean() {
		assetsBean.setLiquidAssets(liquidBox.getValue());
		assetsBean.setFundInvestment(fundBox.getValue());
		assetsBean.setFixedAssets(fixedBox.getValue());
		assetsBean.setIntangibleAsset(intangibleBox.getValue());
		assetsBean.setOtherAssets(otherBox.getValue());
		
	}
}
