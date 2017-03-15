package com.alm.wrapper.classes.ui;

import java.awt.Dimension;
import java.awt.GridLayout;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JPasswordField;
import javax.swing.JTextField;
import javax.swing.UIManager;
import javax.swing.UIManager.LookAndFeelInfo;
import javax.swing.UnsupportedLookAndFeelException;

import com.alm.wrapper.classes.jacob.ALMAutomationWrapper;
import com.alm.wrapper.model.data.ALMData;
import com.alm.wrapper.model.enums.Domains;
import com.alm.wrapper.model.enums.Projects;
import com.jacob.com.LibraryLoader;

/**
 * Class representing the ALM Login window of the tool. Contains main method to
 * trigger the execution.
 * 
 * @author sahil.srivastava
 *
 */
public class ALMLoginWindow extends JFrame implements ActionListener {

	private static final long serialVersionUID = 1L;

	private static final String FAILED_CONN_STRING = "Unable to connect to ALM...!!!";
	private static final String SUCCESS_CONN_STRING = "Connected to ALM...";
	private static final String AUTH_FAIL_STRING = "Authentication failed";
	private static final String INVALID_USERN_PASS_STRING = "Invalid username/password";

	private ALMAutomationWrapper almAutomationWrapper;
	private static ALMData almData;

	private String almURLDefault = "ALM URL";

	private JPanel panel;

	// Declaring Labels
	private JLabel almURLLabel;
	private JLabel userNameLabel;
	private JLabel passwordLabel;
	private JLabel domainLabel;
	private JLabel projectLabel;
	private JLabel logLabel;

	// Declaring Text Fields
	private JTextField almURLTextField;
	private JTextField userNameTextField;
	private JPasswordField passwordTextField;

	private JComboBox<String> domainCombo;
	private String[] domainList = { Domains.DOMAIN_NAME1.getDomain(), Domains.DOMAIN_NAME2.getDomain() };

	private JComboBox<String> projectCombo;
	private String[] projectList = { Projects.PROJECT_NAME1.getProject(),
			Projects.PROJECT_NAME2.getProject() };

	// Declaring Buttons
	private JButton loginButton;
	private JButton clearButton;

	private boolean isConnected;

	public ALMLoginWindow() {
		setSize(400, 300);
		setTitle("ALM Login Details");
		setLayout(null);

		setResizable(false);
		Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
		int w = getSize().width;
		int h = getSize().height;
		int x = (dim.width - w) / 2;
		int y = (dim.height - h) / 2;

		setLocation(x, y);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		panel = new JPanel(new GridLayout(7, 2));
		panel.setLayout(null);
		setContentPane(panel);

		// Adding almURLLabel
		almURLLabel = new JLabel("ALM URL");
		almURLLabel.setLabelFor(almURLTextField);
		almURLLabel.setBounds(50, 25, 100, 25);
		panel.add(almURLLabel);

		// Adding almURLTextField
		almURLTextField = new JTextField(getAlmURLDefault(), 20);
		almURLTextField.setBounds(150, 25, 220, 23);
		almURLTextField.addKeyListener(new KeyListenerImplementation());
		panel.add(almURLTextField);

		// Adding userNameLabel
		userNameLabel = new JLabel("User Name");
		userNameLabel.setLabelFor(userNameTextField);
		userNameLabel.setBounds(50, 55, 100, 25);
		panel.add(userNameLabel);

		// Adding userNameTextField
		userNameTextField = new JTextField(10);
		userNameTextField.setBounds(150, 55, 120, 23);
		userNameTextField.addKeyListener(new KeyListenerImplementation());
		panel.add(userNameTextField);

		// Adding passwordLabel
		passwordLabel = new JLabel("Password");
		passwordLabel.setLabelFor(passwordTextField);
		passwordLabel.setBounds(50, 85, 100, 25);
		panel.add(passwordLabel);

		// Adding passwordTextField
		passwordTextField = new JPasswordField(10);
		passwordTextField.setBounds(150, 85, 120, 25);
		passwordTextField.addKeyListener(new KeyListenerImplementation());
		panel.add(passwordTextField);

		// Adding domainLabel
		domainLabel = new JLabel("Domain");
		domainLabel.setLabelFor(domainCombo);
		domainLabel.setBounds(50, 115, 100, 25);
		panel.add(domainLabel);

		// Adding domainCombo
		domainCombo = new JComboBox<String>(domainList);
		domainCombo.setBounds(150, 115, 220, 25);
		domainCombo.setEditable(true);
		panel.add(domainCombo);

		// Adding projectLabel
		projectLabel = new JLabel("Project");
		projectLabel.setLabelFor(projectCombo);
		projectLabel.setBounds(50, 145, 100, 25);
		panel.add(projectLabel);

		// Adding projectCombo
		projectCombo = new JComboBox<String>(projectList);
		projectCombo.setBounds(150, 145, 220, 25);
		projectCombo.setEditable(true);
		panel.add(projectCombo);

		// Adding log Label
		logLabel = new JLabel();
		logLabel.setBounds(120, 180, 200, 25);
		logLabel.setHorizontalAlignment((int) CENTER_ALIGNMENT);
		panel.add(logLabel);

		// Adding loginButton
		loginButton = new JButton("Login");
		loginButton.setBounds(70, 215, 100, 25);
		loginButton.addActionListener(this);
		loginButton.addKeyListener(new KeyListenerImplementation());
		panel.add(loginButton);

		// Adding clearButton
		clearButton = new JButton("Clear");
		clearButton.setBounds(200, 215, 100, 25);
		clearButton.addActionListener(this);
		clearButton.addKeyListener(new KeyListenerImplementation());
		panel.add(clearButton);

		setVisible(true);

	}

	public String getAlmURLDefault() {
		return almURLDefault;
	}

	public void setAlmURLDefault(String almURL) {
		this.almURLDefault = almURL;
	}

	@Override
	public void actionPerformed(ActionEvent event) {
		if (event.getSource() == loginButton) {
			logLabel.setText("Logging in...");
			logLabel.paintImmediately(logLabel.getVisibleRect());
			ALMData.setAlmURL(almURLTextField.getText());
			ALMData.setUserName(userNameTextField.getText());
			ALMData.setPassword(passwordTextField.getPassword());
			ALMData.setDomain(domainCombo.getSelectedItem().toString());
			ALMData.setProject(projectCombo.getSelectedItem().toString());

			almAutomationWrapper = new ALMAutomationWrapper(almData);
			System.out.println("ALM automation wrapper object created");

			try {
				isConnected = almAutomationWrapper.connectAndLoginALM();
				if (!isConnected) {
					logLabel.setText(FAILED_CONN_STRING);
				} else {
					System.out.println(SUCCESS_CONN_STRING);
					setVisible(false);
					new ALMTestExecutionWindow(almAutomationWrapper);
				}
			} catch (Exception e) {
				System.err.println(e.getMessage());
				if (e.getMessage().contains(AUTH_FAIL_STRING)) {
					logLabel.setText(INVALID_USERN_PASS_STRING);
				} else {
					logLabel.setText(FAILED_CONN_STRING);
				}
				almAutomationWrapper.closeConnection();
			}
		}

		if (event.getSource() == clearButton) {
			userNameTextField.setText("");
			passwordTextField.setText("");
		}
	}

	public static void main(String... s) {
		System.setProperty("jacob.dll.path", System.getProperty("user.dir") + "\\jacob-1.18-x86.dll");
		LibraryLoader.loadJacobLibrary();

		// Changing the Look and Feel of the UI to Nimbus
		try {
			for (LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
				if ("Nimbus".equals(info.getName())) {
					UIManager.setLookAndFeel(info.getClassName());
					break;
				}
			}
		} catch (Exception e) {
			try {
				UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
			} catch (ClassNotFoundException | InstantiationException | IllegalAccessException
					| UnsupportedLookAndFeelException e1) {
				e1.printStackTrace();
			}
		}
		new ALMLoginWindow();
		almData = new ALMData();
	}

	/**
	 * Implementation of KeyListener interface to invoke action events when
	 * specific keys are pressed
	 * 
	 * @author sahil.srivastava
	 *
	 */
	public class KeyListenerImplementation implements KeyListener {

		@Override
		public void keyTyped(KeyEvent arg0) {
			// Definition not required as of now
		}

		@Override
		public void keyReleased(KeyEvent arg0) {
			// Definition not required as of now
		}

		@Override
		public void keyPressed(KeyEvent event) {
			// If the focus is on clear button and enter key is pressed
			if (event.getSource() == clearButton && event.getKeyCode() == KeyEvent.VK_ENTER) {
				actionPerformed(new ActionEvent(clearButton, 2, ""));
			}
			// If enter key is pressed
			else if (event.getKeyCode() == KeyEvent.VK_ENTER) {
				actionPerformed(new ActionEvent(loginButton, 1, ""));
			}
		}
	}
}
