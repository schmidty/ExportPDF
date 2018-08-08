import java.text.*;
import java.io.*;
import java.util.*;
import java.util.regex.*;
import java.util.regex.Pattern;


public class ReadPDFWriteCSV {
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
	
	File csvFile;
	BufferedWriter out;

	public static void main(String[] args) //{{{
	{
		ReadPDFWriteCSV t = new ReadPDFWriteCSV(args[0], args[1]);
	}//}}}
	
	public ReadPDFWriteCSV(String readFileName, String writeFileName) //{{{
	{

		// Name file
		boolean fileExists = false;
		String writeFileNameCSV = writeFileName + ".csv";	

	 	csvFile = new File(writeFileNameCSV);

		fileExists = csvFile.exists();
		if(fileExists) {
			System.out.println("File exists!");
			try {
				FileWriter writeFile = new FileWriter(writeFileNameCSV, true);
			 	out = new BufferedWriter(writeFile);
			} catch (Exception e) {
				e.printStackTrace();
			} 
		} else {
			System.out.println("File does not exist!");
			try {
				FileWriter writeFile = new FileWriter(writeFileNameCSV);
				out = new BufferedWriter(writeFile);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}

		// Now extract text from PDF doc and write out each line
		try {
			// Read the pdf doc
			PDDocument pd = PDDocument.load(readFileName);
			PDFTextStripper f = new PDFTextStripper();

			// Split each line
			String data = f.getText(pd);
			String lines[] = data.split("\\r?\\n");

			for(int x=0; x<lines.length; x++) {
				String aline = lines[x];
				String tempLine = aline.trim();
				//System.out.println(tempLine);

				// Write each line
				out.write(tempLine + "\n");
			}
			pd.close();
			out.close();
			System.out.println("WriteFile closed...");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}//}}}
	
}
