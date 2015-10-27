/* ColorScaleImpl.java

	Purpose:
		
	Description:
		
	History:
		Oct 26, 2015 2:12:50 PM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.util.ArrayList;
import java.util.List;

import org.zkoss.zss.model.SCFValueObject;
import org.zkoss.zss.model.SColor;
import org.zkoss.zss.model.SColorScale;

/**
 * 
 * @author henri
 * @since 3.8.2
 */
public class ColorScaleImpl implements SColorScale {
	private List<SCFValueObject> valueObjects;
	private List<SColor> colors;
	
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
	public List<SColor> getColors() {
		return colors;
	}
	
	public void addColor(SColor color) {
		if (colors == null) {
			colors = new ArrayList<SColor>();
		}
		colors.add(color);
	}

}
