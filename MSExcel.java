import java.io.*;
import java.io.File;
import java.util.Date;

public class MSExcel {
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
	String fileName;
	WritableWorkbook workBook;
	WritableSheet SHEET;
	WritableSheet Sheet;
	Workbook readWorkBook = null;
	static int rowCnt = 0;

	public void WriteWorkBook() //{{{
	{
		try {
			workBook.write();
		} catch(Exception e) {
			e.printStackTrace();
			System.exit(0);
		}
	}//}}}

	public void CloseWorkBook() //{{{
	{
		try {
			workBook.close();
		} catch(Exception e) {
			e.printStackTrace();
			System.exit(0);
		}
	}//}}}

	public MSExcel(String fileName) //{{{
	{
		try {
			workBook = Workbook.createWorkbook(new File(fileName + ".xls"));
			SHEET = workBook.createSheet("PATIENT DATA", 0);
		} catch(Exception e) {
			e.printStackTrace();
			System.exit(1);
		}
	}//}}}

	public void AddHeader() //{{{
	{
			// Create top row labels
			try {
				WritableFont boldFont = new WritableFont(WritableFont.COURIER, 10, WritableFont.BOLD);
				WritableCellFormat boldFontHeader = new WritableCellFormat(boldFont);

				Label lblFirstName = new Label(0,0,"FIRST_NAME", boldFontHeader);		
				SHEET.setColumnView(0, 13);
				SHEET.addCell(lblFirstName);
				Label lblLastName = new Label(1,0,"LAST_NAME", boldFontHeader);
				SHEET.setColumnView(1, 12);
				SHEET.addCell(lblLastName);
				Label lblChartNum = new Label(2,0,"CHART_NUMBER", boldFontHeader);
				SHEET.setColumnView(2, 18);
				SHEET.addCell(lblChartNum);
				Label lblAddress = new Label(3,0,"ADDRESS", boldFontHeader);
				SHEET.setColumnView(3, 25);
				SHEET.addCell(lblAddress);
				Label lblCity = new Label(4,0,"CITY", boldFontHeader);
				SHEET.setColumnView(4, 26);
				SHEET.addCell(lblCity);
				Label lblStateZip = new Label(5,0, "STATE_ZIP", boldFontHeader);
				SHEET.setColumnView(5, 20);
				SHEET.addCell(lblStateZip);
				Label lblHomePhone = new Label(6,0, "HOME_PHONE", boldFontHeader);
				SHEET.setColumnView(6, 27);
				SHEET.addCell(lblHomePhone);
				Label lblWorkPhone = new Label(7,0, "WORK_PHONE", boldFontHeader);
				SHEET.setColumnView(7, 27);
				SHEET.addCell(lblWorkPhone);
				Label lblDoctor = new Label(8,0, "DOCTOR", boldFontHeader);
				SHEET.setColumnView(8, 15);
				SHEET.addCell(lblDoctor);
			} catch(Exception e) {
				e.printStackTrace();
				System.exit(1);	
			}
	}//}}}

	public void AddData(int col, int row, String val) //{{{
	{
		try {
			Label cellVal = new Label(col, row, val);
			SHEET.addCell(cellVal);
		} catch (Exception e) {
			e.printStackTrace();
			System.exit(1);
		}
	}//}}}

	public void FreezeTopRow(int theRowToFreeze) //{{{
	{
		try {
			SHEET.getSettings().setVerticalFreeze(theRowToFreeze);
		} catch(Exception e) {
			e.printStackTrace();
			System.exit(1);
		}
	}//}}}
	

}
