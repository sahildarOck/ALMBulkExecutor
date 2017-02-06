package com.alm.wrapper.classes;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.TimeZone;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.alm.wrapper.enums.AttachmentFor;
import com.alm.wrapper.enums.ExecutionLevel;
import com.alm.wrapper.exceptions.ExcelParsingException;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComFailException;
import com.jacob.com.Dispatch;
import com.jacob.com.EnumVariant;
import com.jacob.com.Variant;

import atu.alm.wrapper.classes.TDConnection;
import atu.alm.wrapper.enums.StatusAs;
import jxl.read.biff.BiffException;

/**
 * Class to wrap the functionalities pertaining to ALM - OTA COM API
 * 
 * @author sahil.srivastava
 *
 */

public class ALMAutomationWrapper {

	// private static ALMServiceWrapper wrapper;
	private TDConnection tdConn;
	private ActiveXComponent almActiveXComponent;
	private ALMData almData;
	private ActiveXComponent defectObject = null;

	private final String DEFAULT_ACTUAL_RESULT = "As expected";
	private final String DEFAULT_BLOCKED_FAILED_RESULT = "Blocked/Failed due to defect: ";

	private String almURL;
	private String userName;
	private String password;
	private String domain;
	private String project;
	private String runName;

	private ExecutionLevel executionFor;

	public ALMAutomationWrapper(ALMData almData) {
		this.almData = almData;

		almURL = ALMData.getAlmURL();
		userName = ALMData.getUserName();
		password = String.valueOf(ALMData.getPassword());
		domain = ALMData.getDomain();
		project = ALMData.getProject();
	}

	/*
	 * public boolean connectAndLoginALM() throws ALMServiceException {
	 * setALMServiceWrapper(new ALMServiceWrapper(almURL));
	 * System.out.println("Inside connectAndLoginALM method"); return
	 * wrapper.connect(userName, password, domain, project); }
	 */

	public boolean connectAndLoginALM() {
		almActiveXComponent = new ActiveXComponent("TDAPIOLE80.TDConnection");
		almActiveXComponent.invoke("InitConnectionEx", almURL);
		almActiveXComponent.invoke("Login", new Variant(userName), new Variant(password));
		almActiveXComponent.invoke("Connect", new Variant(domain), new Variant(project));
		return true;
	}

	public void closeConnection() {
		// wrapper.close();
		System.out.println("invoking log out");
		almActiveXComponent.invoke("Logout");
	}

	/*
	 * public static ALMServiceWrapper getALMServiceWrapper() { return wrapper;
	 * }
	 * 
	 * public static void setALMServiceWrapper(ALMServiceWrapper wrapper) {
	 * ALMAutomationWrapper.wrapper = wrapper; }
	 */

	/*
	 * public static void setALMServiceWrapper(ALMServiceWrapper wrapper) {
	 * ALMAutomationWrapper.wrapper = wrapper; }
	 */

	public ALMData getAlmData() {
		return almData;
	}

	public void setAlmData(ALMData almData) {
		this.almData = almData;
	}

	public ExecutionLevel getExecutionFor() {
		return executionFor;
	}

	public void setExecutionFor(ExecutionLevel executionFor) {
		this.executionFor = executionFor;
	}

	public TDConnection getTdConn() {
		return tdConn;
	}

	public void setTdConn(TDConnection tdConn) {
		this.tdConn = tdConn;
	}

	public ActiveXComponent getAlmActiveXComponent() {
		return almActiveXComponent;
	}

	public void setAlmActiveXComponent(ActiveXComponent almActiveXComponent) {
		this.almActiveXComponent = almActiveXComponent;
	}

	public void executeInALM(ALMTestExecutionWindow testExecWindow)
			throws BiffException, IOException, ExcelParsingException, InvalidFormatException {

		testExecWindow.updateLogLabelForExecution("Preparing execution...");

		// Test Set Default Execution
		if (almData.getTestCaseName().equals("")) {
			setExecutionFor(ExecutionLevel.TEST_SET_DEFAULT);
			testSetDefaultExecution(testExecWindow);
		}

		// Test Case Default Execution - No need to upload test steps
		// excel file; Execution performed on existing design steps with
		// DEFAULT_ACTUAL_RESULT for passed test steps.
		else if (!almData.getTestCaseName().equals("") && almData.getTestStepsXLFile() == null) {
			setExecutionFor(ExecutionLevel.TEST_CASE_DEFAULT);
			testCaseDefaultExecution(testExecWindow);
		}

		// Test Case Customized Execution - Need to upload test steps
		// excel file; Execution performed based on test steps in excel file;
		// Can add a field/column 'Actual Result' beside 'Expected Result'; For
		// blank Actual Results, DEFAULT_ACTUAL_RESULT will be provided
		else {
			setExecutionFor(ExecutionLevel.TEST_CASE_CUSTOMIZED);
			testCaseCustomizedExecution(testExecWindow);
		}

	}

	/**
	 * Test Set Default Execution: Method to block/unblock the tests present in
	 * the Test Set. Only executes if testCaseStatus provided by user is either
	 * Blocked or No run. If testCaseStatus provided is No run by user, it sets
	 * all the blocked tests in the test set to No run and vice-versa if the
	 * testCaseStatus provided by user is Blocked. Also links defects if
	 * defectID provided in the UI
	 */
	private void testSetDefaultExecution(ALMTestExecutionWindow testExecWindow) {
		if (almData.getTestCaseStatus() == StatusAs.BLOCKED || almData.getTestCaseStatus() == StatusAs.NO_RUN) {
			int testCaseNo = 1;

			// Using the below variable as defect is getting linked before
			// iterating for the first test case

			testExecWindow.updateLogLabelForExecution("Executing for each test case of the test set");
			// wrapper.connect(userName, password, domain, project);

			// tdConn = wrapper.getAlmObj();
			// almActiveXComponent = tdConn.getAlmObject();

			// Getting testSetFactory object
			ActiveXComponent testSetFactory = almActiveXComponent.getPropertyAsComponent("TestSetFactory");

			// Getting the specific testSet object
			ActiveXComponent testSet = testSetFactory.invokeGetComponent("Item", new Variant(almData.getTestSetID()));

			// Getting the testCaseFactory object pertaining to the testSet
			ActiveXComponent testCaseFactory = testSet.getPropertyAsComponent("TSTestFactory");

			// Fetching all the testCases in testCaseVariant
			Variant testCasesVariant = testCaseFactory.invoke("NewList", "");

			Dispatch tsTests = testCasesVariant.getDispatch();

			System.out.println("Total no. of test cases: " + Dispatch.get(tsTests, "count"));

			EnumVariant enumVariantTSTests = new EnumVariant(tsTests);

			Dispatch dispatchTSTest;
			ActiveXComponent activeXTSTest = null;

			// Setting up the defect object
			if (almData.getTestCaseStatus() == StatusAs.BLOCKED && !almData.getDefectID().equals("")) {
				defectObject = getDefectObject();
			}

			enumVariantTSTests = new EnumVariant(tsTests);

			// Iterating through the test set to change the execution status
			while (enumVariantTSTests.hasMoreElements()) {
				dispatchTSTest = enumVariantTSTests.nextElement().getDispatch();
				activeXTSTest = new ActiveXComponent(dispatchTSTest);

				// Condition 1: Unblocking the test cases
				if (activeXTSTest.getPropertyAsString("Status").equals("Blocked")
						&& almData.getTestCaseStatus() == StatusAs.NO_RUN) {
					activeXTSTest.setProperty("Status", "No Run");
					activeXTSTest.invoke("Post");
					testExecWindow.updateLogLabelForExecution("Status updated for test case : " + testCaseNo++);
					System.out.println(
							"Status updated for Test Case name: " + activeXTSTest.getPropertyAsString("TestName"));
				}

				// Condition 2: Blocking and linking defects (if applicable) to
				// test cases
				else if (activeXTSTest.getPropertyAsString("Status").equals("No Run")
						&& almData.getTestCaseStatus() == StatusAs.BLOCKED) {
					ActiveXComponent runFactory = activeXTSTest.getPropertyAsComponent("RunFactory");
					ActiveXComponent run = runFactory.invokeGetComponent("AddItem", new Variant(createAndGetRunName()));
					run.invoke("CopyDesignSteps");
					ActiveXComponent stepFactory = run.invokeGetComponent("StepFactory");
					Variant stepsVariant = stepFactory.invoke("NewList", "");
					EnumVariant enumVariantSteps = new EnumVariant(stepsVariant.getDispatch());
					updateStepStatusAndActualResult(StatusAs.BLOCKED,
							new ActiveXComponent(enumVariantSteps.nextElement().getDispatch()),
							DEFAULT_BLOCKED_FAILED_RESULT + almData.getDefectID());
					updateRunStatus(StatusAs.BLOCKED, run);
					run.invoke("Post");

					// Defect linking if defect id is provided
					if (defectObject != null) {
						try {
							ActiveXComponent bugLinkFactory = activeXTSTest.getPropertyAsComponent("BugLinkFactory");
							bugLinkFactory.invoke("AddItem", new Variant(defectObject));
						} catch (ComFailException e) {
							System.err.println("Defect ID doesn't exist..!!");
							testExecWindow.updateLogLabelForExecution("Defect ID doesn't exist..!!");
							defectObject = null;
						}
					}
					activeXTSTest.invoke("Post");
					testExecWindow.updateLogLabelForExecution("Status updated for test case : " + testCaseNo++);
					System.out.println(
							"Status updated for Test Case name: " + activeXTSTest.getPropertyAsString("TestName"));
				}
			}
			testExecWindow.updateLogLabelForExecution("Status updated for " + (testCaseNo - 1) + " test cases");
		} else {
			testExecWindow.updateLogLabelForExecution(
					"Can't set status for the entire tests in Test Set as: " + almData.getTestCaseStatus());
		}
	}

	/**
	 * Test Case Default Execution - Method to perform execution on existing
	 * design steps with DEFAULT_ACTUAL_RESULT for passed test steps. No need to
	 * upload test steps excel file
	 * 
	 * @param testExecWindow
	 */
	private void testCaseDefaultExecution(ALMTestExecutionWindow testExecWindow) {
		System.out.println("Inside testCaseDefaultExecution method");

		// Getting testSetFactory object
		ActiveXComponent testSetFactory = almActiveXComponent.getPropertyAsComponent("TestSetFactory");

		// Getting the specific testSet object
		ActiveXComponent tsTestSet = testSetFactory.invokeGetComponent("Item", new Variant(almData.getTestSetID()));

		// Getting the testCaseFactory object pertaining to the testSet
		ActiveXComponent tsTestFactory = tsTestSet.getPropertyAsComponent("TSTestFactory");

		// Fetching all the testCases in testCaseVariant
		Variant testCasesVariant = tsTestFactory.invoke("NewList", "");

		EnumVariant enumVariant = new EnumVariant(testCasesVariant.getDispatch());

		ActiveXComponent tsTest = null;

		// Fetching the required testCase
		boolean foundTestCase = false;
		while (enumVariant.hasMoreElements()) {
			tsTest = new ActiveXComponent(enumVariant.nextElement().getDispatch());
			if (tsTest.getPropertyAsString("TestName").equals(almData.getTestCaseName())) {
				foundTestCase = true;
				break;
			}
		}

		// Executing test case if it has been found in the previous step
		if (!foundTestCase) {
			testExecWindow.updateLogLabelForExecution("Given test case not found...");
			System.out.println("Given test case not found...");
		} else {
			if (almData.getTestCaseStatus() == StatusAs.BLOCKED) {
				tsTest.setProperty("Status", "Blocked");
				tsTest.invoke("Post");
				if (almData.getAttachmentFile() != null) {
					testExecWindow.updateLogLabelForExecution("Attaching snapshot...");
					attachSnapshot(tsTest);
					testExecWindow.updateLogLabelForExecution("Snapshot attached and test case blocked...");
				} else {
					testExecWindow.updateLogLabelForExecution("Test case blocked...");
				}
			} else if (almData.getTestCaseStatus() == StatusAs.NO_RUN) {
				tsTest.setProperty("Status", "No Run");
				tsTest.invoke("Post");
				testExecWindow.updateLogLabelForExecution("Test case unblocked...");

			} else {
				ActiveXComponent runFactory = tsTest.getPropertyAsComponent("RunFactory");
				ActiveXComponent run = runFactory.invokeGetComponent("AddItem", new Variant(createAndGetRunName()));
				run.invoke("CopyDesignSteps");
				ActiveXComponent stepFactory = run.invokeGetComponent("StepFactory");
				Variant stepsVariant = stepFactory.invoke("NewList", "");
				EnumVariant enumVariantSteps = new EnumVariant(stepsVariant.getDispatch());

				int stepNo = 1;
				if (almData.getTestCaseStatus() == StatusAs.PASSED) {
					System.out.println("Inside passed");
					while (enumVariantSteps.hasMoreElements()) {
						updateStepStatusAndActualResult(StatusAs.PASSED,
								new ActiveXComponent(enumVariantSteps.nextElement().getDispatch()),
								DEFAULT_ACTUAL_RESULT);
						testExecWindow.updateLogLabelForExecution("Updated step No.:" + stepNo);
						System.out.println("Updated step No.:" + stepNo);
						stepNo++;
					}
					updateRunStatus(StatusAs.PASSED, run);
				} else {
					System.out.println("Inside failed");
					while (stepNo < Integer.parseInt(almData.getFailedTestStepNo())) {
						updateStepStatusAndActualResult(StatusAs.PASSED,
								new ActiveXComponent(enumVariantSteps.nextElement().getDispatch()),
								DEFAULT_ACTUAL_RESULT);
						testExecWindow.updateLogLabelForExecution("Updated step No.:" + stepNo);
						System.out.println("Updated step No.:" + stepNo);
						stepNo++;
					}
					updateStepStatusAndActualResult(StatusAs.FAILED,
							new ActiveXComponent(enumVariantSteps.nextElement().getDispatch()),
							almData.getFailedTestStepActualResult());
					testExecWindow.updateLogLabelForExecution("Status updated for " + (stepNo - 1) + " steps");
					updateRunStatus(StatusAs.FAILED, run);
				}
				if (almData.getAttachmentFile() != null) {
					System.out.println("Status updated for " + stepNo + " steps. Attaching snapshot...");
					testExecWindow.updateLogLabelForExecution(
							"Status updated for " + stepNo + " steps. Attaching snapshot...");
					if (almData.getAttachmentFor() == AttachmentFor.RUN) {
						attachSnapshot(run);
					} else {
						attachSnapshot(tsTest);
					}
					System.out.println("Status updated for " + stepNo + " steps and snapshots attached");
					testExecWindow.updateLogLabelForExecution(
							"Status updated for " + stepNo + " steps and snapshots attached");
				} else {
					System.out.println("Status updated for " + (stepNo - 1) + " steps");
				}
				run.invoke("Post");
			}
		}
	}

	/**
	 * TODO: Implement this. Test Case Customized Execution - Method to perform
	 * execution based on test steps in excel file; Can add a field/column
	 * 'Actual Result' beside 'Expected Result'; For blank Actual Results,
	 * DEFAULT_ACTUAL_RESULT will be provided; Need to upload test steps excel
	 * file
	 * 
	 * @param testExecWindow
	 * @throws InvalidFormatException
	 * @throws IOException
	 * @throws ExcelParsingException
	 */
	private void testCaseCustomizedExecution(ALMTestExecutionWindow testExecWindow)
			throws InvalidFormatException, IOException, ExcelParsingException {
		/*
		 * ALMExcelHandler xlHandle = new
		 * ALMExcelHandler(almData.getXlFile().getAbsolutePath()); List<ALMStep>
		 * almSteps = xlHandle.getALMSteps();
		 * 
		 * wrapper.connect(userName, password, domain, project);
		 * 
		 * TestSet testSet = (TestSet) wrapper.getTestSet("", "",
		 * almData.getTestSetID()); TSTestFactory testFactory =
		 * testSet.getTSTestFactory(); ListWrapper<TSTest> testCaseList =
		 * testFactory.getNewList(); ITestCase testCase = null;
		 * 
		 * for (TSTest testCaseItr : testCaseList) { if
		 * (testCaseItr.getTestName().equals(almData.getTestCaseName())) {
		 * testCase = testCaseItr; } }
		 * 
		 * ITestCaseRun testRun = wrapper.createNewRun(testCase,
		 * createAndGetRunName(), almData.getTestCaseStatus());
		 * 
		 * // Test Case status : Passed if
		 * (almData.getTestCaseStatus().equals(StatusAs.PASSED)) { int stepNo =
		 * 1; for (ALMStep almStep : almSteps) {
		 * System.out.println("Executing Step: " + stepNo);
		 * wrapper.addStep(testRun, almStep.getStepName(), StatusAs.PASSED,
		 * almStep.getStepDescription(), almStep.getStepExpectedResult(),
		 * EXPECTED_MSG);
		 * testExecWindow.updateLogLabelForExecution("Executing Step No.: " +
		 * stepNo++); }
		 * testExecWindow.updateLogLabelForExecution("Execution complete..."); }
		 * 
		 * // Test Case status : Failed if
		 * (almData.getTestCaseStatus().equals(StatusAs.FAILED)) { int stepNo =
		 * 1; int failedStepNo =
		 * Integer.parseInt(almData.getFailedTestStepNo()); for (ALMStep almStep
		 * : almSteps) { System.out.println("Executing Step: " + stepNo); if
		 * (stepNo < failedStepNo) { wrapper.addStep(testRun,
		 * almStep.getStepName(), StatusAs.PASSED, almStep.getStepDescription(),
		 * almStep.getStepExpectedResult(), EXPECTED_MSG); } else if (stepNo ==
		 * failedStepNo) { wrapper.addStep(testRun, almStep.getStepName(),
		 * StatusAs.FAILED, almStep.getStepDescription(),
		 * almStep.getStepExpectedResult(),
		 * almData.getFailedTestStepActualResult()); } else {
		 * wrapper.addStep(testRun, almStep.getStepName(), StatusAs.NO_RUN,
		 * almStep.getStepDescription(), almStep.getStepExpectedResult(), ""); }
		 * testExecWindow.updateLogLabelForExecution("Executing Step No.: " +
		 * stepNo++); }
		 * testExecWindow.updateLogLabelForExecution("Execution complete..."); }
		 * 
		 * if (almData.getAttachmentFile() != null) {
		 * testExecWindow.updateLogLabelForExecution("Attaching snaps..."); if
		 * (almData.getAttachmentFor().equals(AttachmentFor.TESTCASE)) {
		 * wrapper.newAttachment(almData.getAttachmentFile().getAbsolutePath(),
		 * "testCaseAttachemnt", testCase); testExecWindow.
		 * updateLogLabelForExecution("Snaps attached and execution complete..."
		 * ); } else {
		 * wrapper.newAttachment(almData.getAttachmentFile().getAbsolutePath(),
		 * "testRunAttachemnt", testRun); testExecWindow.
		 * updateLogLabelForExecution("Snaps attached and execution complete..."
		 * ); } }
		 */

		// Getting testSetFactory object
		ActiveXComponent testSetFactory = almActiveXComponent.getPropertyAsComponent("TestSetFactory");

		// Getting the specific testSet object
		ActiveXComponent tsTestSet = testSetFactory.invokeGetComponent("Item", new Variant(almData.getTestSetID()));

		// Getting the testCaseFactory object pertaining to the testSet
		ActiveXComponent tsTestFactory = tsTestSet.getPropertyAsComponent("TSTestFactory");

		// Fetching all the testCases in testCaseVariant
		Variant testCasesVariant = tsTestFactory.invoke("NewList", "");

		EnumVariant enumVariant = new EnumVariant(testCasesVariant.getDispatch());

		ActiveXComponent tsTest = null;

		// Fetching the required testCase
		boolean foundTestCase = false;
		while (enumVariant.hasMoreElements()) {
			tsTest = new ActiveXComponent(enumVariant.nextElement().getDispatch());
			if (tsTest.getPropertyAsString("TestName").equals(almData.getTestCaseName())) {
				foundTestCase = true;
				break;
			}
		}

		// Executing test case if it has been found in the previous step
		if (!foundTestCase) {
			testExecWindow.updateLogLabelForExecution("Given test case not found...");
			System.out.println("Given test case not found...");
		} else {
			ActiveXComponent runFactory = tsTest.getPropertyAsComponent("RunFactory");
			ActiveXComponent run = runFactory.invokeGetComponent("AddItem", new Variant(createAndGetRunName()));
			run.invoke("CopyDesignSteps");
			ActiveXComponent stepFactory = run.invokeGetComponent("StepFactory");
			Variant stepsVariant = stepFactory.invoke("NewList", "");
			EnumVariant enumVariantSteps = new EnumVariant(stepsVariant.getDispatch());

			while (enumVariantSteps.hasMoreElements()) {
				ActiveXComponent step = new ActiveXComponent(enumVariantSteps.nextElement().getDispatch());
				step.setProperty("Status", "Passed");
				setProperty(step, "Field", new Variant[] { new Variant("ST_ACTUAL") },
						new Variant("Testing the actual result"));
				step.invoke("Post");
			}
		}
	}

	public void setProperty(Dispatch activex, String name, Variant[] indexes, Variant value) {
		Variant[] variants = new Variant[indexes.length + 1];

		for (int i = 0; i < indexes.length; i++) {
			variants[i] = indexes[i];
		}
		variants[variants.length - 1] = value;
		Dispatch.invoke(activex, name, Dispatch.Put, variants, new int[variants.length]);
	}

	/**
	 * Method to attach snapshot
	 * 
	 * @param tsTestOrRun
	 * @param testExecWindow
	 */
	private boolean attachSnapshot(ActiveXComponent tsTestOrRun) {
		ActiveXComponent attachmentFactory = tsTestOrRun.getPropertyAsComponent("Attachments");
		Variant variantAttachment = new Variant();
		variantAttachment.putNull();
		ActiveXComponent attachment = attachmentFactory.invokeGetComponent("AddItem", variantAttachment);
		attachment.setProperty("Type", new Variant(1L));
		attachment.setProperty("FileName", almData.getAttachmentFile().getAbsolutePath());
		attachment.invoke("Post");
		return true;
	}

	/**
	 * Method to create and get the run name in the specific format as it is
	 * created while running manually Format: 'Run_M-d_h-m-s'
	 * 
	 * @return - runName in format 'Run_M-d_h-m-s'
	 */
	private String createAndGetRunName() {
		SimpleDateFormat dateFormat = new SimpleDateFormat("M-d_h-m-s");
		dateFormat.setTimeZone(TimeZone.getTimeZone("GMT-04:00"));
		runName = "Run_" + dateFormat.format(new Date());
		getAlmData().setRunName(runName);
		return runName;
	}

	private ActiveXComponent getDefectObject() {
		// tdConn = wrapper.getAlmObj();
		// almActiveXComponent = tdConn.getAlmObject();

		ActiveXComponent bugFactory = almActiveXComponent.getPropertyAsComponent("BugFactory");
		return bugFactory.invokeGetComponent("Item", new Variant(almData.getDefectID()));
	}

	/**
	 * Method to update the test step's status and update the actual result
	 * 
	 * @param step
	 *            - ActiveXComponent
	 */
	private void updateStepStatusAndActualResult(StatusAs status, ActiveXComponent step, String actualResult) {
		if (status.equals(StatusAs.PASSED)) {
			step.setProperty("Status", "Passed");
			System.out.println("passed the step");
		} else if (status.equals(StatusAs.FAILED)) {
			step.setProperty("Status", "Failed");
		}
		setProperty(step, "Field", new Variant[] { new Variant("ST_ACTUAL") }, new Variant(actualResult));
		step.invoke("Post");
		System.out.println("posted");
	}

	private void updateRunStatus(StatusAs status, ActiveXComponent run) {
		if (status.equals(StatusAs.PASSED)) {
			run.setProperty("Status", "Passed");
		} else if (status.equals(StatusAs.FAILED)) {
			run.setProperty("Status", "Failed");
		}
	}
}