package com.alm.wrapper.model.data;

import atu.alm.wrapper.enums.StatusAs;

/**
 * @deprecated Class representing a single ALM Step. Was required in earlier
 *             version of the tool. Might be of use in further versions when
 *             adding the feature of adding customized actual results for the
 *             test cases
 * @author sahil.srivastava
 *
 */
public class ALMStep {

	private String stepName;
	private String stepDescription;
	private String stepExpectedResult;
	private String stepActualResult;
	private StatusAs stepStatus;

	public String getStepName() {
		return stepName;
	}

	public void setStepName(String stepName) {
		this.stepName = stepName;
	}

	public String getStepDescription() {
		return stepDescription;
	}

	public void setStepDescription(String stepDescription) {
		this.stepDescription = stepDescription;
	}

	public String getStepExpectedResult() {
		return stepExpectedResult;
	}

	public void setStepExpectedResult(String expectedResult) {
		this.stepExpectedResult = expectedResult;
	}

	public String getStepActualResult() {
		return stepActualResult;
	}

	public void setStepActualResult(String actualResult) {
		this.stepActualResult = actualResult;
	}

	public StatusAs getStepStatus() {
		return stepStatus;
	}

	public void setStepStatus(StatusAs stepStatus) {
		this.stepStatus = stepStatus;
	}
}
