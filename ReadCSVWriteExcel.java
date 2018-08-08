import java.io.*;
import java.util.*;
import java.util.regex.*;
import java.util.regex.Pattern;

public class ReadCSVWriteExcel {
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

	WritableWorkbook workbook;
	String firstName;
	String lastName;
	String doctor;
	String street;
	String apt;
	String city;
	String stateZip;
	String homePhone;
	String workPhone;
	String chartNumber;

	public static void main(String[] args) 
	{
		ReadCSVWriteExcel d = new ReadCSVWriteExcel(args[0]);
	}

	public ReadCSVWriteExcel(String openFile) //{{{
	{
		String strLine;
		int lineCnt = 1;
		int x = 0;
		Vector<String> chartNums = new Vector<String>();
		boolean writeData = true;

		// NameCase each value
		NameCase n = new NameCase();

		// Replace the ".csv"
		String csvEndingString = ".csv$";
		String replaceEndingString = "";
		Pattern csvEndingPattern = Pattern.compile(csvEndingString);
		Matcher matcher = csvEndingPattern.matcher(openFile);

		String writeFile = matcher.replaceAll(replaceEndingString);
		System.out.println("ExcelFileName: " + writeFile);

		// Instantiate the MSExcel class
		MSExcel d = new MSExcel(writeFile);
		d.AddHeader();

		// Regular Expressions compile
		String drNameStringBy = "Appointment Book contains '(.*?)'";
		Pattern drNameStringByPattern = Pattern.compile(drNameStringBy);

		String  wholeNameString = "(.*?[,{1}].*?)0\\d{8}";
		Pattern wholeNamePattern = Pattern.compile(wholeNameString);

		String  chartNumString = "(0\\d{8})";
		Pattern chartNumPattern = Pattern.compile(chartNumString);

		String  addressString = ",.*,\\s\\D{2}\\s";
		Pattern addressPattern = Pattern.compile(addressString);

		String  phoneString = "Phone";
		Pattern phoneStringPattern = Pattern.compile(phoneString);
		
		try {
			// Open csv file for reading
			BufferedReader br = new BufferedReader(new FileReader(openFile));

			// Go through each line and parse out needed values
			while((strLine = br.readLine()) !=null) {
				// Count each line in the CSV file and report to command-line
				x++;


				Matcher chartNumMatcher = chartNumPattern.matcher(strLine);
				if(chartNumMatcher.find() == true) {
					String s = chartNumMatcher.group(1);
					System.out.println("Line " + x + " -  Chart Number: " + s);

					chartNumber = s;

					// Add to vector charNums in not already in vector chartNums and set writeData to 'true'
					if(chartNums.contains(s)==false) {
						chartNums.add(s);
						writeData = true;
					} else {
						writeData = false;
					}

					if(writeData == true) {
						// Add to data to Excel sheet
						d.AddData(2, lineCnt, chartNumber);
					}
				}
				Matcher drNameMatcher = drNameStringByPattern.matcher(strLine);
				if(drNameMatcher.find() == true && writeData == true) {
					String s = drNameMatcher.group(1);
					System.out.println("Line " + x + " -  Doctor: " + s);
					doctor = NameCase.transformCase(s);

					// Doctor gets written on every line for column 8 (see last line of this loop)
				}
				Matcher wholeNameMatcher = wholeNamePattern.matcher(strLine);
				if(wholeNameMatcher.find() == true && writeData == true) {
					String s = wholeNameMatcher.group(1);
					String wholeNameArray[] = s.split(", ");
					String fNameArray[] = wholeNameArray[1].split(" ");
					String fName = fNameArray[0];
					String lName = wholeNameArray[0];
					//System.out.println("Line " + x + " -  Whole Name: " + s);
					System.out.println("Line " + x + " -  First Name: " + fName);
					System.out.println("Line " + x + " -  Last Name: " + lName);

					firstName = NameCase.transformCase(fName);
					lastName = NameCase.transformCase(lName);

					// Add to data to Excel sheet
					d.AddData(0, lineCnt, firstName);
					d.AddData(1, lineCnt, lastName);
				}
				Matcher addressMatcher = addressPattern.matcher(strLine);
				if(addressMatcher.find() == true && x > 3 && writeData == true) {
					//System.out.println("Line " + x + " -  Address: " + strLine);
					
					String addressParts[] = strLine.split(", ");

					if(addressParts.length == 3) {
						street = NameCase.transformCase(addressParts[0]);
						city = NameCase.transformCase(addressParts[1]);
						stateZip = NameCase.transformCase(addressParts[2]);
						System.out.println("Line " + x + " -  Street: " + street);

						// Add the street to Excel 
						d.AddData(3, lineCnt, street);
					} else 
					if(addressParts.length == 4) {
						street = NameCase.transformCase(addressParts[0]);
						apt = NameCase.transformCase(addressParts[1]);
						city = NameCase.transformCase(addressParts[2]);
						stateZip = NameCase.transformCase(addressParts[3]);

						// split up state and zip code
						System.out.println("Line " + x + " -  Street: " + street);
						System.out.println("Line " + x + " -  Apt: " + apt);

						// Add the street and apt to Excel 
						d.AddData(3, lineCnt, street + " " + apt);
					}

					System.out.println("Line " + x + " -  City: " + city);
					System.out.println("Line " + x + " -  StateZip: " + stateZip);

					// Add to data to Excel sheet
					d.AddData(4, lineCnt, city);
					d.AddData(5, lineCnt, stateZip);
				}
				Matcher phoneMatcher = phoneStringPattern.matcher(strLine);
				if(phoneMatcher.find() == true && writeData == true) {
					//System.out.println("Line " + x + " -  Phones: " + strLine);

					String phoneNums[] = strLine.split("       ");

					System.out.println("Line " + x + " -  " + phoneNums[0]);
					System.out.println("Line " + x + " -  " + phoneNums[1]);

					homePhone = phoneNums[0];
					workPhone = phoneNums[1];

					// Add to data to Excel sheet
					d.AddData(6, lineCnt, homePhone);
					d.AddData(7, lineCnt, workPhone);

					// Seperate the lines
					System.out.println("=====================================");

					// Add the doctor to the last column
					d.AddData(8, lineCnt, doctor);

					// Next line for Excel spreadsheet
					lineCnt++;
				}
				// Freeze the top header row
				d.FreezeTopRow(1);

			}
			// Write and close the Excel workbook
			d.WriteWorkBook();
			d.CloseWorkBook();
			br.close();
		} catch (Exception e) {
			System.out.println("pdfbox error");
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
	}//}}} 

	public void DeleteCSVFile(String fileName) //{{{
	{
		try {
			File deleteTarget = new File(fileName);
			deleteTarget.delete();
		} catch (Exception e) {
			System.out.println("ERROR! Could not delete CSV file!");
			e.printStackTrace();
		}

	}//}}}

}
