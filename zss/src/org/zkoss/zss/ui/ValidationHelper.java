package org.zkoss.zss.ui;

import org.zkoss.poi.ss.usermodel.DataValidation;
import org.zkoss.poi.ss.usermodel.DataValidation.ErrorStyle;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.impl.SheetImpl;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.XRanges;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zul.Messagebox;

/**
 * 
 * @author dennis
 *
 */
//Code refer from Spreadsheet.validate
public class ValidationHelper {
	Spreadsheet ss;
	public ValidationHelper(Spreadsheet ss) {
		this.ss = ss;
	}

	private boolean isEventThreadEnabled() {
		return Executions.getCurrent().getDesktop().getWebApp()
				.getConfiguration().isEventThreadEnabled();
	}

	// return true if a valid input; false otherwise and show Error Alert if
	// required
	public boolean validate(Sheet sheet, final int row, final int col,
			final String editText, final EventListener callback) {
		final Sheet ssheet = ss.getSelectedSheet();
		if (ssheet == null || !ssheet.equals(sheet)) { //skip no sheet case
			return true;
		}
		if (_inCallback) { // skip validation check
			return true;
		}
		final XRange rng = XRanges.range(((SheetImpl)sheet).getNative(), row, col);
		final DataValidation dv = rng.validate(editText);
		if (dv != null) {
			if (dv.getShowErrorBox()) {
				String errTitle = dv.getErrorBoxTitle();
				String errText = dv.getErrorBoxText();
				if (errTitle == null) {
					errTitle = "ZK Spreadsheet";
				}
				if (errText == null) {
					errText = "The value you entered is not valid.\n\nA user has restricted values that can be entered into this cell.";
				}
				final int errStyle = dv.getErrorStyle();
				switch (errStyle) {
				case ErrorStyle.STOP: {
					final int btn = Messagebox.show(errText, errTitle,
							Messagebox.RETRY | Messagebox.CANCEL,
							Messagebox.ERROR, Messagebox.RETRY,
							new EventListener() {
								public void onEvent(Event event)
										throws Exception {
									final String evtname = event.getName();
									if (Messagebox.ON_RETRY.equals(evtname)) {
										retry(callback);
									} else if (Messagebox.ON_CANCEL
											.equals(evtname)) {
										cancel(callback);
									}
								}
							});
				}
					break;
				case ErrorStyle.WARNING: {
					errText += "\n\nContinue?";
					final int btn = Messagebox.show(errText, errTitle,
							Messagebox.YES | Messagebox.NO | Messagebox.CANCEL,
							Messagebox.EXCLAMATION, Messagebox.NO,
							new EventListener() {
								public void onEvent(Event event)
										throws Exception {
									final String evtname = event.getName();
									if (Messagebox.ON_NO.equals(evtname)) {
										retry(callback);
									} else if (Messagebox.ON_CANCEL
											.equals(evtname)) {
										cancel(callback);
									} else if (Messagebox.ON_YES
											.equals(evtname)) {
										ok(callback);
									}
								}
							});
					if (isEventThreadEnabled() && btn == Messagebox.YES) {
						return true;
					}
				}
					break;
				case ErrorStyle.INFO: {
					final int btn = Messagebox.show(errText, errTitle,
							Messagebox.OK | Messagebox.CANCEL,
							Messagebox.INFORMATION, Messagebox.OK,
							new EventListener() {
								@Override
								public void onEvent(Event event)
										throws Exception {
									final String evtname = event.getName();
									if (Messagebox.ON_CANCEL.equals(evtname)) {
										cancel(callback);
									} else if (Messagebox.ON_OK.equals(evtname)) {
										ok(callback);
									}
								}
							});
					if (isEventThreadEnabled() && btn == Messagebox.OK) {
						return true;
					}
				}
					break;
				}
			}
			return false;
		}
		return true;
	}

	private boolean _inCallback = false;

	private void errorBoxCallback(EventListener callback, String eventname) {
		if (!isEventThreadEnabled() && callback != null) {
			try {
				_inCallback = true;
				callback.onEvent(new Event(eventname, ss));
			} catch (Exception e) {
				throw UiException.Aide.wrap(e);
			} finally {
				_inCallback = false;
			}
		}
	}

	// when user press OK/YES button of the validation ErrorBox, have to call
	// back to resend the setEditText() operation
	private void ok(EventListener callback) {
		errorBoxCallback(callback, Messagebox.ON_OK);
	}

	// when user press RETRY/NO button of the validation ErrorBox, have to call
	// back to handle UI operation
	private void retry(EventListener callback) {
		// TODO: shall set focus back to cell at (row, col), select the text,
		// enter edit mode
		errorBoxCallback(callback, Messagebox.ON_RETRY);
	}

	// when user press CANCEL button of the validation ErrorBox, have to call
	// back to handle UI operation
	private void cancel(EventListener callback) {
		// TODO: shall set focus back to cell at (row, col) and restore cell
		// value
		errorBoxCallback(callback, Messagebox.ON_CANCEL);
	}
}
