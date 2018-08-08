import java.io.*;
import java.awt.*;
import java.awt.event.*;
import javax.swing.*;
import javax.swing.SwingUtilities;
import javax.swing.filechooser.FileFilter;
import java.util.*;
import java.util.regex.*;
import java.util.regex.Pattern;
import java.text.*;

public class ExportPDF extends JPanel implements ActionListener {
	/*
	 *		ExportPDF application - Patient data extraction for PDF
	 *		Version 1.3
	 * 
	 *		By Matthew J Schmidt
	 *		Copyright on August 14th, 2011
	 *		All Rights Reserved
	 *		
	 *		No part of this application may be re-used, edited, altered 
	 *		or manipulated without the express written consent of the author
	 *
	 */
	JFileChooser fc;
	static private final String newline = "\n";
	JButton openButton, saveButton, extractButton, closeButton;
	JTextArea log;
	File[] files;
	String pdfFile;
	String pdfFilePath;
	String pdfFileName;

	public static void main(String[] args) //{{{
	{
		SwingUtilities.invokeLater(new Runnable() {
			@Override
			public void run() {
				UIManager.put("swing.boldMetal", Boolean.FALSE);
				createAndShowGUI();
			}
		});
	}//}}}
	
	public static void createAndShowGUI() //{{{
	{
		JFrame frame = new JFrame("ExportPDF - Patient Data Extract By Matt Schmidt Ver 1.3 ");
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setPreferredSize(new Dimension(575, 275));
		frame.setLocation(400, 400);
		frame.add(new ExportPDF());

		frame.pack();
		frame.setVisible(true);
	}//}}}

	public ExportPDF() //{{{
	{
		super(new BorderLayout());

		fc = new JFileChooser();
		log = new JTextArea(20, 20);
		log.setMargin(new Insets(5,5,5,5));
		log.setEditable(false);
		JScrollPane logScrollPane = new JScrollPane(log);

		openButton = new JButton("Open PDF File...", createImageIcon("images/Open16.gif"));
		openButton.addActionListener(this);

		closeButton = new JButton("Close");
		closeButton.addActionListener(this);

		extractButton = new JButton("Extract!");
		extractButton.addActionListener(this);

		JPanel buttonPanel = new JPanel();
		buttonPanel.add(openButton);
		buttonPanel.add(extractButton);
		buttonPanel.add(closeButton);

		add(buttonPanel, BorderLayout.PAGE_START);
		add(logScrollPane, BorderLayout.CENTER);
	}//}}}

	public static ImageIcon createImageIcon(String path) //{{{
	{
		java.net.URL imgURL = ExportPDF.class.getResource(path);
		if(imgURL !=null) {
			return new ImageIcon(imgURL);
		} else {
			System.err.println("Couldn't find image file: " + path);
			return null;
		}
	}//}}}

	@Override
	public void actionPerformed(ActionEvent e) //{{{
	{
		if(e.getSource() == openButton) {
			// Create filter
			JFileChooser chooser = new JFileChooser();
			chooser.setCurrentDirectory(new File("."));
			chooser.setMultiSelectionEnabled(true);
			chooser.setFileFilter(new FileFilter() {
				@Override
				public boolean accept(File f) {
					return f.getName().toLowerCase().endsWith(".pdf") || f.isDirectory();
				}
				@Override
				public String getDescription() {
					return "PDF Files";
				}
			});
			int returnVal = chooser.showOpenDialog(ExportPDF.this);

			if(returnVal == JFileChooser.APPROVE_OPTION) {
				files = chooser.getSelectedFiles();
				for(int x=0; x<files.length; x++) {
					pdfFile = files[x].getName();
					log.append("Opening PDF File: " + pdfFile + " " + newline);
				}
			} else {
				System.out.println("Open command cancelled by user.");
			}
		}
		if(e.getSource() == extractButton) {
			String writeFilePathName = "";
			String writeFilePathNameCSV = "";

			// Get the current date
			Calendar currentDate = Calendar.getInstance();
			SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMdd");
			String dateNow = formatter.format(currentDate.getTime());

			// Extract text from each PDF file
			log.append("Writing CSV file..." + newline);
			log.append("" + newline);
			for(int x=0; x<files.length; x++) {
				pdfFilePath = files[x].getPath();
				pdfFileName = files[x].getName();
				log.append("Extracting text from pdf: " + pdfFileName + "... " + newline);

				// Extract and write the PDF file data to Excel spreadsheet
				if(pdfFile !=null) {
					// Replace the ".pdf" with ".xls" for the Excel spreadsheet
					String writeFileNameString = pdfFileName + "$";
					String writeFileNameReplace = "";
					Pattern writeFileNameReplacePattern = Pattern.compile(writeFileNameString);	
					Matcher writeFileNameReplaceMatcher = writeFileNameReplacePattern.matcher(pdfFilePath);
					String writeFilePath = writeFileNameReplaceMatcher.replaceAll(writeFileNameReplace);
					String writeFileName = pdfFile;
					writeFilePathName = writeFilePath + "PatientData_" + dateNow;
					writeFilePathNameCSV = writeFilePathName + ".csv";

					System.out.println("readPDFFile: " + pdfFilePath);
					System.out.println("writeFilePathName: " + writeFilePathName);
					System.out.println("Writing...");

					
					//ReadPDFWriteExcel rpdf = new ReadPDFWriteExcel(pdfFilePath, writeFilePathName);
					ReadPDFWriteCSV rpdf = new ReadPDFWriteCSV(pdfFilePath, writeFilePathName);
				
					log.append("Extraction of this file done..." + newline);
					log.append("" + newline);
				} else {
					log.append("No file selected!" + newline);
				}
			}
			// Extract from CSV file and write to Excel file
			log.append("Creating Excel file..." + newline);
			ReadCSVWriteExcel writeExcel = new ReadCSVWriteExcel(writeFilePathNameCSV);
			writeExcel.DeleteCSVFile(writeFilePathNameCSV);
			log.append("Excel file complete!" + newline);

			// Close out
			log.append("" + newline);
			log.append("Extraction complete!" + newline);
		}
		if(e.getSource() == closeButton) {
			System.exit(0);
		}
	}//}}}

}

