package com.alm.wrapper.classes;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.alm.wrapper.exceptions.ExcelParsingException;

/**
 * @deprecated Class to handle the excel of format 2000 and below. Was required
 *             in v1.1 of the tool
 * 
 * @author sahil.srivastava
 *
 */
public class ALMExcelHandler {

	private String xlLocation;
	private Workbook xlWorkbook;
	private List<ALMStep> almSteps;

	private final String STEP_NAME = "Step Name";
	private final String DESCRIPTION = "Description";
	private final String EXPECTED_RESULT = "Expected Result";

	public ALMExcelHandler(String xlLocation) throws IOException, InvalidFormatException {
		this.setXlLocation(xlLocation);
		this.setXlWorkbook(new XSSFWorkbook(new File(xlLocation)));
	}

	public ALMExcelHandler(File xlFile) throws IOException, InvalidFormatException {
		this.setXlWorkbook(new XSSFWorkbook(xlFile));
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
		if (xlWorkbook.equals(null)) {
			throw new ExcelParsingException("Excel file not loaded...");
		} else {
			sheet = xlWorkbook.getSheetAt(0);
			Iterator<Row> rowIterator = sheet.iterator();
			Row row1 = rowIterator.next();
			Iterator<Cell> cellIteratorForRow1 = row1.cellIterator();

			if (!(cellIteratorForRow1.next().getStringCellValue().equals(STEP_NAME))) {
				throw new ExcelParsingException("Unable to read excel. A1 cell's content should be 'Step Name'");
			}
			if (!(cellIteratorForRow1.next().getStringCellValue().equals(DESCRIPTION))) {
				throw new ExcelParsingException("Unable to read excel. A2 cell's content should be 'Description'");
			}
			if (!(cellIteratorForRow1.next().getStringCellValue().equals(EXPECTED_RESULT))) {
				throw new ExcelParsingException("Unable to read excel. A3 cell's content should be 'Expected Result'");
			}
			while (rowIterator.hasNext()) {
				Row nextRow = rowIterator.next();
				Iterator<Cell> cellIterator = nextRow.cellIterator();

				ALMStep almStep = new ALMStep();

				String stepName = cellIterator.next().getStringCellValue();
				String stepDescription = cellIterator.next().getStringCellValue();
				String stepExpectedResult = cellIterator.next().getStringCellValue();

				almStep.setStepName(stepName);
				almStep.setStepDescription(stepDescription);
				almStep.setStepExpectedResult(stepExpectedResult);

				almSteps.add(almStep);
			}
		}

		this.setALMSteps(almSteps);
		return almSteps;
	}

	public void setALMSteps(List<ALMStep> almSteps) {
		this.almSteps = almSteps;
	}
}
