package com.alm.wrapper.classes;

import java.awt.Dimension;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;

import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JRadioButton;
import javax.swing.JTextArea;
import javax.swing.JTextField;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.alm.wrapper.enums.AttachmentFor;
import com.alm.wrapper.exceptions.ExcelParsingException;
import com.jacob.activeX.ActiveXComponent;

import atu.alm.wrapper.enums.StatusAs;
import jxl.read.biff.BiffException;

/**
 * Class representing the ALM Test Execution window of the tool
 * 
 * @author sahil.srivastava
 *
 */
public class ALMTestExecutionWindow extends JFrame implements ActionListener {

	private static final long serialVersionUID = 1L;

	private ALMAutomationWrapper almAutomationWrapper;
	private ALMData almData;

	private JPanel panel;

	private JLabel testSetIDLabel;
	private JLabel testCaseNameLabel;
	private JLabel testCaseStatusLabel;
	private JLabel defectIDLabel;
	private JLabel failedTestStepNoLabel;
	private JLabel failedTestStepActualResultLabel;
	private JLabel testStepsXlPathLabel;
	private JLabel attachmentPathLabel;
	private JLabel attachmentForLabel;
	private JLabel logLabel;

	private JTextField testSetIDTextField;
	private JTextField testCaseNameTextField;
	private JTextField defectIDTextField;
	private JTextField failedTestStepNoTextField;
	private JTextField testStepsXlPathTextField;
	private JTextField attachmentPathTextField;

	private JTextArea failedTestStepActualResultTextArea;

	private JComboBox<String> runStatusListCombo;
	// private String[] runStatus = { "Blocked", "Failed", "N/A", "No Run", "Not
	// Completed", "Passed" };
	private String[] runStatus = { "No Run", "Blocked", "Failed", "Passed" };

	private JFileChooser fileChooser;

	private JButton browseXlButton;
	private JButton browseAttachmentButton;
	private JButton executeInAlmButton;

	private ButtonGroup buttonGroup;
	private JRadioButton runRadioButton;
	private JRadioButton testCaseRadioButton;

	public ALMTestExecutionWindow(ALMAutomationWrapper almAutomationWrapper) {

		this.almAutomationWrapper = almAutomationWrapper;
		this.almData = almAutomationWrapper.getAlmData();

		setSize(700, 600);
		setTitle("ALM Test Execution Details ------------- User logged in: " + ALMData.getUserName());
		setLayout(null);

		setResizable(false);
		Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
		int w = getSize().width;
		int h = getSize().height;
		int x = (dim.width - w) / 2;
		int y = (dim.height - h) / 2;

		setLocation(x, y);

		addWindowListener(new java.awt.event.WindowAdapter() {
			@Override
			public void windowClosing(java.awt.event.WindowEvent windowEvent) {
				setVisible(false);
				ActiveXComponent almActiveXComponent = almAutomationWrapper.getAlmActiveXComponent();
				almActiveXComponent.invoke("Disconnect");
				System.out.println("Disconnecting from project...");
				almActiveXComponent.invoke("Logout");
				System.out.println("Terminating the user's connection: logging out...");
				almActiveXComponent.invoke("ReleaseConnection");
				System.out.println("Releasing the COM pointer...");
				System.exit(0);
			}
		});

		panel = new JPanel();
		panel.setLayout(null);
		setContentPane(panel);

		testSetIDLabel = new JLabel("TestSet ID*");
		testSetIDLabel.setLabelFor(testSetIDTextField);
		testSetIDLabel.setBounds(70, 50, 100, 25);
		panel.add(testSetIDLabel);

		testSetIDTextField = new JTextField();
		testSetIDTextField.setBounds(270, 50, 100, 23);
		panel.add(testSetIDTextField);

		testCaseNameLabel = new JLabel("Test Case Name");
		testCaseNameLabel.setLabelFor(testCaseNameTextField);
		testCaseNameLabel.setBounds(70, 80, 100, 25);
		panel.add(testCaseNameLabel);

		testCaseNameTextField = new JTextField();
		testCaseNameTextField.setBounds(270, 80, 300, 23);
		panel.add(testCaseNameTextField);

		testCaseStatusLabel = new JLabel("Test Case Status");
		testCaseStatusLabel.setLabelFor(runStatusListCombo);
		testCaseStatusLabel.setBounds(70, 110, 100, 25);
		panel.add(testCaseStatusLabel);

		runStatusListCombo = new JComboBox<String>(runStatus);
		runStatusListCombo.addActionListener(this);
		runStatusListCombo.setBounds(270, 110, 100, 23);
		panel.add(runStatusListCombo);

		defectIDLabel = new JLabel("Defect ID");
		defectIDLabel.setLabelFor(defectIDTextField);
		defectIDLabel.setBounds(70, 140, 120, 25);
		defectIDLabel.setEnabled(true);
		panel.add(defectIDLabel);

		defectIDTextField = new JTextField();
		defectIDTextField.setBounds(270, 140, 100, 23);
		defectIDTextField.setEnabled(true);
		panel.add(defectIDTextField);

		failedTestStepNoLabel = new JLabel("Failed Test Step no.");
		failedTestStepNoLabel.setLabelFor(failedTestStepNoTextField);
		failedTestStepNoLabel.setBounds(70, 170, 120, 25);
		failedTestStepNoLabel.setEnabled(true);
		panel.add(failedTestStepNoLabel);

		failedTestStepNoTextField = new JTextField();
		failedTestStepNoTextField.setBounds(270, 170, 100, 23);
		failedTestStepNoTextField.setEnabled(false);
		panel.add(failedTestStepNoTextField);

		failedTestStepActualResultLabel = new JLabel("Failed Test Step actual result");
		failedTestStepActualResultLabel.setLabelFor(failedTestStepActualResultTextArea);
		failedTestStepActualResultLabel.setBounds(70, 200, 170, 25);
		failedTestStepActualResultLabel.setEnabled(false);
		panel.add(failedTestStepActualResultLabel);

		failedTestStepActualResultTextArea = new JTextArea();
		failedTestStepActualResultTextArea.setBounds(270, 200, 300, 35);
		failedTestStepActualResultTextArea.setEnabled(false);
		panel.add(failedTestStepActualResultTextArea);

		testStepsXlPathLabel = new JLabel("Test Steps Excel Path");
		testStepsXlPathLabel.setLabelFor(testStepsXlPathTextField);
		testStepsXlPathLabel.setBounds(70, 270, 180, 25);
//		panel.add(testStepsXlPathLabel);

		fileChooser = new JFileChooser();

		testStepsXlPathTextField = new JTextField();
		testStepsXlPathTextField.setBounds(270, 270, 300, 23);
		testStepsXlPathTextField.setEditable(false);
//		panel.add(testStepsXlPathTextField);

		browseXlButton = new JButton("Browse...");
		browseXlButton.setBounds(480, 295, 88, 23);
		browseXlButton.addActionListener(this);
//		panel.add(browseXlButton);

		attachmentPathLabel = new JLabel("Attachment Path");
		attachmentPathLabel.setLabelFor(attachmentPathTextField);
		attachmentPathLabel.setBounds(70, 340, 180, 25);
		panel.add(attachmentPathLabel);

		attachmentPathTextField = new JTextField();
		attachmentPathTextField.setBounds(270, 340, 300, 23);
		attachmentPathTextField.setEditable(false);
		panel.add(attachmentPathTextField);

		browseAttachmentButton = new JButton("Browse...");
		browseAttachmentButton.setBounds(480, 365, 88, 23);
		browseAttachmentButton.addActionListener(this);
		panel.add(browseAttachmentButton);

		attachmentForLabel = new JLabel("Attachment for");
		attachmentForLabel.setBounds(70, 420, 180, 25);
		panel.add(attachmentForLabel);

		runRadioButton = new JRadioButton("Run", true);
		runRadioButton.setBounds(300, 420, 100, 25);
		panel.add(runRadioButton);

		testCaseRadioButton = new JRadioButton("Test Case");
		testCaseRadioButton.setBounds(450, 420, 180, 25);
		panel.add(testCaseRadioButton);

		buttonGroup = new ButtonGroup();
		buttonGroup.add(runRadioButton);
		buttonGroup.add(testCaseRadioButton);

		logLabel = new JLabel("");
		logLabel.setBounds(240, 423, 300, 100);
		logLabel.setHorizontalAlignment((int) CENTER_ALIGNMENT);
		panel.add(logLabel);

		executeInAlmButton = new JButton("Execute In ALM !!!");
		executeInAlmButton.setBounds(250, 510, 180, 25);
		executeInAlmButton.setHorizontalAlignment((int) CENTER_ALIGNMENT);
		executeInAlmButton.addActionListener(this);
		panel.add(executeInAlmButton);

		setVisible(true);
	}

	@Override
	public void actionPerformed(ActionEvent event) {
		if (event.getSource() == runStatusListCombo) {
			/*if (runStatusListCombo.getSelectedItem().equals("Failed")
					|| runStatusListCombo.getSelectedItem().equals("Blocked")) {
				defectIDLabel.setEnabled(true);
				defectIDTextField.setEnabled(true);
			} else {
				defectIDTextField.setText("");
				defectIDLabel.setEnabled(false);
				defectIDTextField.setEnabled(false);
			}*/

			if (runStatusListCombo.getSelectedItem().equals("Failed") && !testCaseNameTextField.getText().equals("")) {
				failedTestStepNoLabel.setEnabled(true);
				failedTestStepNoTextField.setEnabled(true);
				failedTestStepActualResultLabel.setEnabled(true);
				failedTestStepActualResultTextArea.setEnabled(true);
			} else {
				failedTestStepNoTextField.setText("");
				failedTestStepActualResultTextArea.setText("");
				failedTestStepNoLabel.setEnabled(false);
				failedTestStepNoTextField.setEnabled(false);
				failedTestStepActualResultLabel.setEnabled(false);
				failedTestStepActualResultTextArea.setEnabled(false);
			}
		}

		if (event.getSource() == testCaseNameTextField) {
			if (runStatusListCombo.getSelectedItem().equals("Failed") && !testCaseNameTextField.getText().equals("")) {
				failedTestStepNoLabel.setEnabled(true);
				failedTestStepNoTextField.setEnabled(true);
				failedTestStepActualResultLabel.setEnabled(true);
				failedTestStepActualResultTextArea.setEnabled(true);
			} else {
				failedTestStepNoTextField.setText("");
				failedTestStepActualResultTextArea.setText("");
				failedTestStepNoLabel.setEnabled(false);
				failedTestStepNoTextField.setEnabled(false);
				failedTestStepActualResultLabel.setEnabled(false);
				failedTestStepActualResultTextArea.setEnabled(false);
			}
		}

		if (event.getSource() == browseXlButton) {
			int returnVal = fileChooser.showOpenDialog(ALMTestExecutionWindow.this);
			if (returnVal == JFileChooser.APPROVE_OPTION) {
				almData.setTestStepsXLFile(fileChooser.getSelectedFile());
				testStepsXlPathTextField.setText(almData.getTestStepsXLFile().getPath());
			}

		}

		if (event.getSource() == browseAttachmentButton) {
			int returnVal = fileChooser.showOpenDialog(ALMTestExecutionWindow.this);
			if (returnVal == JFileChooser.APPROVE_OPTION) {
				almData.setAttachmentFile(fileChooser.getSelectedFile());
				attachmentPathTextField.setText(almData.getAttachmentFile().getPath());
			}
		}

		if (event.getSource() == executeInAlmButton) {

			if (testSetIDTextField.getText().equals("")) {
				JOptionPane.showMessageDialog(this, "Test Set Id field can't be left blank...!!!", "Warning",
						JOptionPane.WARNING_MESSAGE);
			} else {
				if (!testCaseNameTextField.getText().equals("") && runStatusListCombo.getSelectedItem().equals("Failed")
						&& (failedTestStepNoTextField.getText().equals("")
								|| failedTestStepActualResultTextArea.getText().equals(""))) {
					JOptionPane.showMessageDialog(this,
							"Failed Test Step no./Failed Test Step can't be left blank...!!!", "Warning",
							JOptionPane.WARNING_MESSAGE);
				} else {
					almData.setTestSetID(Integer.parseInt(testSetIDTextField.getText()));
					almData.setTestCaseName(testCaseNameTextField.getText());

					switch (runStatusListCombo.getSelectedItem().toString()) {

					case "Blocked":
						almData.setTestCaseStatus(StatusAs.BLOCKED);
						break;

					case "Failed":
						almData.setTestCaseStatus(StatusAs.FAILED);
						break;

					case "N/A":
						almData.setTestCaseStatus(StatusAs.N_A);
						break;

					case "No Run":
						almData.setTestCaseStatus(StatusAs.NO_RUN);
						break;

					case "Not Completed":
						almData.setTestCaseStatus(StatusAs.NOT_COMPLETED);
						break;

					case "Passed":
						almData.setTestCaseStatus(StatusAs.PASSED);
						break;
					}

					almData.setDefectID(defectIDTextField.getText());
					almData.setFailedTestStepNo(failedTestStepNoTextField.getText());
					almData.setFailedTestStepActualResult(failedTestStepActualResultTextArea.getText());

					if (testCaseRadioButton.isSelected()) {
						almData.setAttachmentFor(AttachmentFor.TESTCASE);
					}

					else if (runRadioButton.isSelected()) {
						almData.setAttachmentFor(AttachmentFor.RUN);
					}

					try {

						// Automation method call
						almAutomationWrapper.executeInALM(this);
					} catch (BiffException e) {
						logLabel.setText(e.getMessage());
					} catch (IOException e) {
						logLabel.setText(e.getMessage());
					} catch (ExcelParsingException e) {
						logLabel.setText(e.getMessage());
					} catch (InvalidFormatException e) {
						logLabel.setText(e.getMessage());
					}
				}
			}
		}
	}

	public void updateLogLabelForExecution(String str) {
		logLabel.setText(str);
		logLabel.paintImmediately(logLabel.getVisibleRect());
	}
}