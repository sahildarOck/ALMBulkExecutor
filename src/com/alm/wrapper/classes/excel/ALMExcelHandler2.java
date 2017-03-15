package com.alm.wrapper.classes.excel;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.alm.wrapper.model.data.ALMStep;
import com.alm.wrapper.model.exceptions.ExcelParsingException;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**
 * @deprecated Class to handle excel of any format. Was required in previous
 *             versions and will be used in further versions when adding feature
 *             of adding customized actual results for the test case execution
 * 
 * 
 * @author sahil.srivastava
 *
 */
public class ALMExcelHandler2 {

	private String xlLocation;
	private Workbook xlWorkbook;
	private List<ALMStep> almSteps;

	public ALMExcelHandler2(String xlLocation) throws BiffException, IOException {
		this.setXlLocation(xlLocation);
		this.setXlWorkbook(Workbook.getWorkbook(new File(xlLocation)));
	}

	public ALMExcelHandler2(File xlFile) throws BiffException, IOException {
		this.setXlWorkbook(Workbook.getWorkbook(xlFile));
	}

	public Workbook getXlWorkbook() {
		return xlWorkbook;
	}

	public void setXlWorkbook(Workbook xlWorkbook) {
		this.xlWorkbook = xlWorkbook;
	}

	public String getXlLocation() {
		return xlLocation;
	}

	public void setXlLocation(String xlLocation) {
		this.xlLocation = xlLocation;
	}

	public List<ALMStep> getALMSteps() throws ExcelParsingException {
		almSteps = new ArrayList<ALMStep>();
		Sheet sheet;
		int stepNo;
		if (xlWorkbook.equals(null)) {
			System.out.println("Excel file not loaded...");
		} else {
			sheet = xlWorkbook.getSheet(0);
			if (!(sheet.getCell(0, 0).getContents().equals("Step Name"))) {
				throw new ExcelParsingException("Unable to read excel. A1 cell's content should be 'Step Name'");
			}
			if (!(sheet.getCell(1, 0).getContents().equals("Description"))) {
				throw new ExcelParsingException("Unable to read excel. A2 cell's content should be 'Description'"
						+ sheet.getCell(0, 1).getContents());
			}
			if (!(sheet.getCell(2, 0).getContents().equals("Expected Result"))) {
				throw new ExcelParsingException("Unable to read excel. A3 cell's content should be 'Expected Result'");
			}
			stepNo = 1;
			try {
				while (!(sheet.getCell(0, stepNo).getContents().equals(""))) {
					ALMStep almStep = new ALMStep();
					String stepName = sheet.getCell(0, stepNo).getContents();
					String stepDescription = sheet.getCell(1, stepNo).getContents();
					String stepExpectedResult = sheet.getCell(2, stepNo).getContents();
					almStep.setStepName(stepName);
					almStep.setStepDescription(stepDescription);
					almStep.setStepExpectedResult(stepExpectedResult);
					almSteps.add(almStep);
					stepNo++;
				}
				System.out.println("Excel read successful...");
			} catch (ArrayIndexOutOfBoundsException e) {
				System.out.println("Excel read successful... ");
			}
		}
		this.setALMSteps(almSteps);
		return almSteps;
	}

	public void setALMSteps(List<ALMStep> almSteps) {
		this.almSteps = almSteps;
	}
}
