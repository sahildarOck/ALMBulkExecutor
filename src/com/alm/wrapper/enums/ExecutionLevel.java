package com.alm.wrapper.enums;

/**
 * Enumeration to identify the level of execution - TestSet: for each test case
 * in the test set, TestCase: for the test case and not for each test step,
 * TestStep: for each test step and the test case
 * 
 * @author sahil.srivastava
 *
 */

public enum ExecutionLevel {
	TEST_SET_DEFAULT,
	TEST_CASE_DEFAULT,
	TEST_CASE_CUSTOMIZED
}
