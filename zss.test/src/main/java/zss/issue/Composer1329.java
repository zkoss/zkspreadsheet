package zss.issue;

import javax.servlet.http.HttpSession;

import org.zkoss.zk.ui.Sessions;
import org.zkoss.zk.ui.util.Composer;
import org.zkoss.zss.ui.Spreadsheet;
//import org.zkoss.zk.ui.Component;

public class Composer1329 implements Composer<Spreadsheet> {

	@Override
	public void doAfterCompose(Spreadsheet comp) throws Exception {
	
		final HttpSession sess = (HttpSession)Sessions.getCurrent().getNativeSession();
		
		Thread simulatedBackgroundOperation = new Thread() {
			public void run() {
				sess.invalidate();
				System.out.println("session invalidated");
			}
		};
		simulatedBackgroundOperation.start();
		
		//wait for background operation to finish 
		//simulating a longer initialization process during which the server closes the HttpSession -> cleaning up the current desktop
		simulatedBackgroundOperation.join();

		//try to access the internal book
		comp.getBook();
		System.out.println("***1329-stack-overflow.zul: It is expected if later you see (HTTP ERROR 500) java.lang.NullPointerException caused by org.zkoss.zk.ui.sys.HtmlPageenders.outSpecialJS(HtmlPageRenders.java:851) ...");
	}

}
