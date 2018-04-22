import java.util.*;
import java.io.*; 
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * 
 * READ ME
 * 
 * Project Specifications:
 * 		Takes in a user specified .xls file, a sheet within the file
 * 		Concatenates Strings in given rows in an Excel .xls sheet and 
 * 		deletes the text in the copied rows, replacing the first
 * 		copied cell's text with the concatenated version. 
 * 		Continues process until user inputs "Q" for quit 
 * 
 * **/

public class RowEditor {
	public static void main(String[] args) throws IOException {
		// takes in the user's .xls file 
		Scanner console = new Scanner(System.in);
		System.out.print("Enter a .xls file path: ");
		String file = console.next();
		
		// create Workbook and load file 
		FileInputStream fis = new FileInputStream(new File(file)); 
		HSSFWorkbook wb = new HSSFWorkbook(fis);
		
		// lists the sheets in the Workbook 
		System.out.println("Sheets:");
		Set<String> wbSheets = new HashSet<String>();
		for (int i = 0; i < wb.getNumberOfSheets(); i++) {
			String sheetName = wb.getSheetName(i);
			wbSheets.add(sheetName);
			System.out.println("    " + sheetName);
		}
		
		// takes in the user's sheet. if the input is not a sheet in the file
		// throws IllegalArgumentException 
		System.out.print("Enter a sheet name: ");
		String sheetName = console.next();
		if (!wbSheets.contains(sheetName)) {
			throw new IllegalArgumentException();
		}
		Sheet sheet = wb.getSheet(sheetName);
		
		// continues asking the user whether they want to continue editing 
		// the rows in their sheet
		boolean quit = false;
		while (!quit) {
			// asks user which rows and columns to concatenate
			int numRows = sheet.getLastRowNum();
			System.out.print("Enter the start index (row 1 = 0): ");
			int startIndex = console.nextInt();
			System.out.print("Enter the end index (row 1 = 0): ");
			int endIndex = console.nextInt();
			System.out.print("Enter a column number (column 1 = 0): "); 
			int colIndex = console.nextInt();
			
			// if the row range is not valid, throws IllegalArgumentException
			if (startIndex > numRows || endIndex > numRows || startIndex >= endIndex) {
				throw new IllegalArgumentException();
			}
			
			// create the concatenated result 
			String result = "";
			for (int i = startIndex; i <= endIndex; i++) {
				Row row = sheet.getRow(i);
				result += " " + row.getCell(colIndex).getStringCellValue();
				row.getCell(colIndex).setCellValue("");
			}
			
			// sets the value of the first cell to the concatenated version 
			Cell cell = sheet.getRow(startIndex).getCell(colIndex);
			cell.setCellValue(result);
			printToExcel(wb, file);
			
			// asks the user if they want to quit 
			System.out.print("Type \"q\" to end or \"n\" to continue: ");
			quit = (console.next().equalsIgnoreCase("q"));
		}
	}
	
	// prints the changes to the .xls file 
	public static void printToExcel(Workbook wb, String file) throws IOException {
		try {
			FileOutputStream output = new FileOutputStream(new File(file));
			wb.write(output);
			output.close();
		} catch(Exception e) {
			e.printStackTrace();
		} 
	}
}
