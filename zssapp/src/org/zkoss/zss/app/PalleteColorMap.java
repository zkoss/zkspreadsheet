package org.zkoss.zss.app;

import java.util.HashMap;
import java.util.Map;

public class PalleteColorMap {
	static Map mapColorHex = new HashMap();

	static {
		mapColorHex = new HashMap();
		mapColorHex.put("AQUA", "#33CCCC");
		mapColorHex.put("AUTOMATIC", "#FFFFFF");
		mapColorHex.put("BLACK", "#000000");
		mapColorHex.put("BLUE", "#0000FF");
		mapColorHex.put("BLUE_GREY", "#6666CC");
		mapColorHex.put("BLUE2", "#0000FF");
		mapColorHex.put("BRIGHT_GREEN", "#00FF00");
		mapColorHex.put("BROWN", "#993300");
		mapColorHex.put("CORAL", "#FF8080");
		mapColorHex.put("DARK_BLUE", "#000080");
		mapColorHex.put("DARK_BLUE2", "#000080");
		mapColorHex.put("DARK_GREEN", "#003300");
		mapColorHex.put("DARK_PURPLE", "#660066");
		mapColorHex.put("DARK_RED", "#800000");
		mapColorHex.put("DARK_RED2", "#800000");
		mapColorHex.put("DARK_TEAL", "#003366");
		mapColorHex.put("DARK_YELLOW", "#808000");
		mapColorHex.put("DEFAULT_BACKGROUND", "#FFFFFF");
		mapColorHex.put("DEFAULT_BACKGROUND1", "#FFFFF");
		mapColorHex.put("GOLD", "#FFCC00");
		mapColorHex.put("GRAY_25", "#C0C0C0");
		mapColorHex.put("GRAY_50", "#808080");
		mapColorHex.put("GRAY_80", "#333333");
		mapColorHex.put("GREEN", "#008000");
		mapColorHex.put("GREY_15_PERCENT", "#D0D0D0");
		mapColorHex.put("GREY_25_PERCENT", "#C0C0C0");
		mapColorHex.put("GREY_40_PERCENT", "#969696");
		mapColorHex.put("GREY_50_PERCENT", "#808080");
		mapColorHex.put("GREY_80_PERCENT", "#333333");
		mapColorHex.put("ICE_BLUE", "#CCCCFF");
		mapColorHex.put("INDIGO", "#333399");
		mapColorHex.put("IVORY", "#FFFFCC");
		mapColorHex.put("LAVENDER", "#CC99FF");
		mapColorHex.put("LIGHT_BLUE", "#3366FF");
		mapColorHex.put("LIGHT_GREEN", "#CCFFCC");
		mapColorHex.put("LIGHT_ORANGE", "#FF9900");
		mapColorHex.put("LIGHT_TURQUOISE", "#CCFFFF");
		mapColorHex.put("LIGHT_TURQUOISE2", "#CCFFFF");
		mapColorHex.put("LIME", "#99CC00");
		mapColorHex.put("OCEAN_BLUE", "#0066CC");
		mapColorHex.put("OLIVE_GREEN", "#333300");
		mapColorHex.put("ORANGE", "#FF6600");
		mapColorHex.put("PALE_BLUE", "#99CCFF");
		mapColorHex.put("PALETTE_BLACK", "#010000");
		mapColorHex.put("PERIWINKLE", "#9999FFF");
		mapColorHex.put("PINK", "#FF00FF");
		mapColorHex.put("PINK2", "#FF00FF");
		mapColorHex.put("PLUM", "#993366");
		mapColorHex.put("PLUM2", "#993366");
		mapColorHex.put("RED", "#FF0000");
		mapColorHex.put("ROSE", "#FF99CC");
		mapColorHex.put("SEA_GREEN", "#339966");
		mapColorHex.put("SKY_BLUE", "#00CCFF");
		mapColorHex.put("TAN", "#FFCC99");
		mapColorHex.put("TEAL", "#008080");
		mapColorHex.put("TEAL2", "#008080");
		mapColorHex.put("TURQOISE2", "#00FFFF");
		mapColorHex.put("TURQUOISE", "#00FFFF");
		mapColorHex.put("UNKNOWN", "#000000");
		mapColorHex.put("VERY_LIGHT_YELLOW", "#FFCC99");
		mapColorHex.put("VIOLET", "#808000");
		mapColorHex.put("VIOLET2", "#800080");
		mapColorHex.put("WHITE", "#FFFFFF");
		mapColorHex.put("YELLOW", "#FFFF00");
		mapColorHex.put("YELLOW2", "#FFFF00");
	}
	
	static String getColorHex(String color){
		return (String) mapColorHex.get(color);
	}
}
