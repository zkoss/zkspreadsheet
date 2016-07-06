/* IconSetImpl.java

	Purpose:
		
	Description:
		
	History:
		Oct 26, 2015 2:23:13 PM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.zkoss.zss.model.SCFValueObject;
import org.zkoss.zss.model.SColor;
import org.zkoss.zss.model.SIconSet;

/**
 * @author henri
 * @since 3.8.2
 */
public class IconSetImpl implements SIconSet, Serializable {
	private static final long serialVersionUID = 167905784918493054L;
	//ZSS-1142
	private static final Map<String, String[]> iconSets = new HashMap<String, String[]>(17);
	static {
		iconSets.put("3ArrowsGray", new String[] {"arrowsGrayDown", "arrowsGrayRight", "arrowsGrayUp"});
		iconSets.put("3Flags", new String[] {"flagsRed", "flagsYellow", "flagsGreen"});
		iconSets.put("3TrafficLights1", new String[] {"trafficLights1Red", "trafficLights1Yellow", "trafficLights1Green"}); //default set
		iconSets.put("3TrafficLights2", new String[] {"trafficLights2Red", "trafficLights2Yellow", "trafficLights2Green"});
		iconSets.put("3Signs", new String[] {"signsRed", "signsYellow", "signsGreen"});
		iconSets.put("3Symbols", new String[] {"symbolsRed", "symbolsYellow", "symbolsGreen"});
		iconSets.put("3Symbols2", new String[] {"symbols2Red", "symbols2Yellow", "symbols2Green"});
		iconSets.put("3Arrows", new String[] {"arrowsDown", "arrowsRight", "arrowsUp"});
		iconSets.put("4Arrows", new String[] {"arrowsDown", "arrowsDownRight", "arrowsUpRight", "arrowsUp"});
		iconSets.put("4ArrowsGray", new String[] {"arrowsGrayDown", "arrowsGrayDownRight", "arrowsGrayUpRight", "arrowsGrayUp"});
		iconSets.put("4RedToBlack", new String[] {"trafficLightsBlack", "trafficLightsGray", "trafficLightsPink", "trafficLights1Red"});
		iconSets.put("4Rating", new String[] {"signal1", "signal2", "signal3", "signal4"});
		iconSets.put("4TrafficLights", new String[] {"trafficLightsBlack", "trafficLights1Red", "trafficLights1Yellow", "trafficLights1Green"});
		iconSets.put("5Arrows", new String[] {"arrowsDown", "arrowsDownRight", "arrowsRight", "arrowsUpRight", "arrowsUp"});
		iconSets.put("5ArrowsGray", new String[] {"arrowsGrayDown", "arrowsGrayDownRight", "arrowsGrayRight", "arrowsGrayUpRight", "arrowsGrayUp"});
		iconSets.put("5Rating", new String[] {"signal0", "signal1", "signal2", "signal3", "signal4"});
		iconSets.put("5Quarters", new String[] {"quarters0", "quarters1", "quarters2", "quarters3", "quarters4"});
		
		// Excel 2016
		iconSets.put("3Stars", new String[] {"start0", "start1", "start2"});
		iconSets.put("3Triangles", new String[] {"triangle1", "triangle2", "triangle3"});
		iconSets.put("5Boxes", new String[] {"boxes0", "boxes1", "boxes2", "boxes3", "boxes4"});
	}
	
	public static String getIconSetName(String name, int iconSetId, boolean reverse) {
		final String[] sets = iconSets.get(name);
		final int iconSetId0 = reverse ?  (sets.length - iconSetId - 1) : iconSetId; 
		return sets == null ? null : sets[iconSetId0];
	}
	
	private IconSetType type;
	private List<SCFValueObject> valueObjects;
	private boolean percent;
	private boolean reverse;
	private boolean showValue = true; //ZSS-1161: default true
	
	
	@Override
	public IconSetType getType() {
		return type;
	}
	
	public void setType(IconSetType type) {
		this.type = type;
	}

	@Override
	public List<SCFValueObject> getCFValueObjects() {
		return valueObjects;
	}
	
	public void addValueObject(SCFValueObject vo) {
		if (valueObjects == null) {
			valueObjects = new ArrayList<SCFValueObject>();
		}
		valueObjects.add(vo);
	}

	@Override
	public boolean isPercent() {
		return percent;
	}
	
	public void setPercent(boolean b) {
		percent = b;
	}

	@Override
	public boolean isReverse() {
		return reverse;
	}

	public void setReverse(boolean b) {
		reverse = b;
	}
	
	@Override
	public boolean isShowValue() {
		return showValue;
	}
	
	public void setShowValue(boolean b) {
		showValue = b;
	}

	//ZSS-1142
	public IconSetImpl cloneIconSet() {
		IconSetImpl iconSet = new IconSetImpl();
		
		iconSet.type = type;
		iconSet.percent = percent;
		iconSet.reverse = reverse;
		iconSet.showValue = iconSet.showValue;
		for (SCFValueObject vo : valueObjects) {
			CFValueObjectImpl vo0 = (CFValueObjectImpl) vo;
			iconSet.addValueObject(vo0.cloneCFValueObject());
		}

		return iconSet;
	}
}
