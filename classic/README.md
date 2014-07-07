Classic Theme
========

ZK Spreadsheet 3.5.0 Theme

Usage:

1. Put classic.jar in WEB-INF/lib, then Classic will become your default theme if there is no other themes.
2. Now you can also dynamically switch between different themes by cookie or library property
	* Use library-property in in WEB-INF/zk.xml

			<library-property>
				<name>org.zkoss.zss.theme.preferred</name>
				<value>classic</value>
			</library-property>
	* Use cookie to switch theme, add a cookie

			zsstheme=classic

It does not require a server restart, but user has to refresh the browser.