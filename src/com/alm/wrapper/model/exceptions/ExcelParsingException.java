package com.alm.wrapper.model.exceptions;

/**
 * @deprecated Customized exception while excel handling. Might be of use in
 *             further versions of the tool
 * @author sahil.srivastava
 *
 */
public class ExcelParsingException extends Exception {

	private static final long serialVersionUID = 1L;

	public ExcelParsingException(String errorMsg) {
		super(errorMsg);
	}
}
