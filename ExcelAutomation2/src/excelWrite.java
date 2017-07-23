
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collections;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelWrite {
	
	static String[] columnsNeeded = {"TAM", "SSR", "Organization Name", "Contract Title", 
										"Contract Number", "Schedule Name", "Value", "TPID",
										"Service Name", "Start Date", "End Date"};
	
	static ArrayList<String> columnsNeeded_ = new ArrayList<String> (
			Arrays.asList("TAM", "SSR", "Organization Name", "Contract Title", 
						"Contract Number", "Schedule Name", "Value", "TPID",
						"Service Name", "Start Date", "End Date"));
	
	static String[] miaColumns = {"TAM", "SISR", "Customer", "Contract Title",
									"Contract", "Schedule Name", "RM Value", "TPID",
									"Service", "Start Date", "End Date"};
	
	static String[] HTAMs = {"abusmt", "alea", "azrinaki", "clam",
							"colee", "easonlau", "taufiqo", "tuchong"};
	
	static String file1;
	static String file2;
	static String outputFile;
			
	public static void main(String[] args) throws IOException {
		
		ArrayList<String> excelFiles = getAllExcel();
		checkExcel(excelFiles);
		
		//Check if output file is already opened
		//to be added
		
		try {
			String filepath1 = excelFiles.get(0);
			String filepath2 = excelFiles.get(1);
			
			FileInputStream file1 = new FileInputStream(new File(filepath1));
			FileInputStream file2 = new FileInputStream(new File(filepath2));
			
			long file1Size = file1.getChannel().size();
			long file2Size = file2.getChannel().size();
			
			//Check which file size is bigger. Bigger file size is Unicorn.
			if(file1Size > file2Size) {
				execute(file1, file2);
			}
			else {
				execute(file2, file1);
			}
			
			file1.close();
			file2.close();
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}
	
	public static void execute_() throws IOException {
		ArrayList<String> excelFiles = getAllExcel();
		checkExcel(excelFiles);
		
		//Check if output file is already opened
		//to be added
		
		try {
			String filepath1 = excelFiles.get(0);
			String filepath2 = excelFiles.get(1);
			
			FileInputStream file1 = new FileInputStream(new File(filepath1));
			FileInputStream file2 = new FileInputStream(new File(filepath2));
			
			long file1Size = file1.getChannel().size();
			long file2Size = file2.getChannel().size();
			
			//Check which file size is bigger. Bigger file size is Unicorn.
			if(file1Size > file2Size) {
				execute(file1, file2);
			}
			else {
				execute(file2, file1);
			}
			
			file1.close();
			file2.close();
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}
	
	public static void execute(FileInputStream file1, FileInputStream file2) {
		ArrayList<ArrayList<String>> excelSheet1 = new ArrayList<ArrayList<String>>();
		ArrayList<ArrayList<String>> excelSheet2 = new ArrayList<ArrayList<String>>();
		
		String sheetName1 = "Sheet1";
		String sheetName2 = "Detailed Schedule";
		
		read(file1, excelSheet1, sheetName1);
		
		//check excel sheet
		ArrayList<ArrayList<String>> output = filter(excelSheet1);
		
		//Beginning of first processing
		int rowSize = output.size();
		int columnSize = output.get(0).size();
		
		sortColumns(output);
		renameColumns(output);
		
		for(int i = 1; i < rowSize; i++) {
			formatDate_(output.get(i), columnSize - 1);
			correctPrice(output.get(i));
			convertLowerCase(output.get(i));
		}
		
		//Beginning of second processing
		ArrayList<ArrayList<XSSFCellStyle>> style = new ArrayList<ArrayList<XSSFCellStyle>>();
		
		read2(file2, excelSheet2, sheetName2, style);
		ArrayList<ArrayList<String>> output2 = filter2(excelSheet2);
		ArrayList<ArrayList<XSSFCellStyle>> styleOutput = filterStyle(style);
		
		int rowSize2 = output2.size();
		int columnSize2 = output2.get(0).size();
		
		for(int i = 1; i < rowSize2; i++) {
			formatDate_(output2.get(i), columnSize2 - 1);
		}
		
		addRows(output, output2);

		//write
		
	}
		
	public static void read(FileInputStream file, ArrayList<ArrayList<String>> excelSheet, String sheetName) {
		try {
	        //Create Workbook instance holding reference to .xlsx file
	        XSSFWorkbook workbook = new XSSFWorkbook(file);
	        
	        workbook.setMissingCellPolicy(MissingCellPolicy.RETURN_BLANK_AS_NULL);
	        DataFormatter fmt = new DataFormatter();
	        
	        XSSFSheet sheet;
	        if(sheetName.equals("null")) {
	        	sheet = workbook.getSheetAt(0);
	        }
	        else {
	        	sheet = workbook.getSheet(sheetName);
	        }
        	
        	for(int rn = sheet.getFirstRowNum(); rn <= sheet.getLastRowNum(); rn++) {
        		//Create 2nd-dimension array list
	            ArrayList<String> columns = new ArrayList<String>();
	            
        		Row row = sheet.getRow(rn);
        		if(row == null) {
        			//There is no data in this row. Need to handle appropriately
        			columns.add("null");
        		}
        		else {
        			for(int cn = 0; cn < row.getLastCellNum(); cn++) {
        				Cell cell = row.getCell(cn);
        				if(cell == null) {
        					//This cell is empty/blank/unused, handle appropriately
        					columns.add("null");
        				}
        				else {
        					String cellStr = fmt.formatCellValue(cell);
        					//Do something with the value
        					columns.add(cellStr);
        				}
        			}
        		}
        		excelSheet.add(columns);
        	}
        	workbook.close();
		}
		catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	//Same as read, tailored to Mia's excel sheet. Will copy the style and formatting.
	public static void read2(FileInputStream file, ArrayList<ArrayList<String>> excelSheet, String sheetName,
			ArrayList<ArrayList<XSSFCellStyle>> style) {
		try {
	        //Create Workbook instance holding reference to .xlsx file
	        XSSFWorkbook workbook = new XSSFWorkbook(file);
	        
	        workbook.setMissingCellPolicy(MissingCellPolicy.RETURN_BLANK_AS_NULL);
	        DataFormatter fmt = new DataFormatter();
	        
	        XSSFSheet sheet;
	        if(sheetName.equals("null")) {
	        	sheet = workbook.getSheetAt(0);
	        }
	        else {
	        	sheet = workbook.getSheet(sheetName);
	        }
	        
	        ArrayList<XSSFCellStyle> styleRow = new ArrayList<XSSFCellStyle>();
        	
        	for(int rn = sheet.getFirstRowNum(); rn <= sheet.getLastRowNum(); rn++) {
        		//Create 2nd-dimension array list
	            ArrayList<String> columns = new ArrayList<String>();
	            
        		XSSFRow row = sheet.getRow(rn);
        		if(row == null) {
        			//There is no data in this row. Need to handle appropriately
        			columns.add("null");
        		}
        		else {
        			for(int cn = 0; cn < row.getLastCellNum(); cn++) {
        				XSSFCell cell = row.getCell(cn);
        				if(cell == null) {
        					//This cell is empty/blank/unused, handle appropriately
        					columns.add("null");
        				}
        				else {
        					String cellStr = fmt.formatCellValue(cell);
        					//Do something with the value
        					columns.add(cellStr);
        					
        					XSSFCellStyle s = cell.getCellStyle();
        					styleRow.add(s);
        				}
        			}
        		}
        		excelSheet.add(columns);
        		style.add(styleRow);
        	}
        	workbook.close();
		}
		catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public static void write(FileInputStream file, ArrayList<ArrayList<String>> excelSheet, String sheetName, 
			String filename, String filepath2, ArrayList<ArrayList<XSSFCellStyle>> style) {
		try {
			//Create Workbook instance holding reference to .xlsx file
	        XSSFWorkbook workbook = new XSSFWorkbook(file);
	        
	        
		} 
		catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	//Traverse through the columns, get the indexes of the columns needed
	public static ArrayList<Integer> checkColumns(ArrayList<ArrayList<String>> array) {
		ArrayList<Integer> columnNumbers = new ArrayList<Integer>();
		
		int columnSize = array.get(0).size();
		
		for(int i = 0; i < columnSize; i++) {
			if(Arrays.asList(columnsNeeded).contains(array.get(0).get(i))) {
				columnNumbers.add(i);
			}
		}
		
		return columnNumbers;
	}
	
	//Traverse through the rows, get the indexes of the rows needed
	public static ArrayList<Integer> checkRows(ArrayList<ArrayList<String>> array) {
		ArrayList<Integer> rowNumbers = new ArrayList<Integer>();
		
		int rowSize = array.size();
		
		for(int i = 0; i < rowSize; i++) {
			if(Arrays.asList(HTAMs).contains(array.get(i).get(0))) {
				rowNumbers.add(i);
			}
		}
		
		return rowNumbers;
	}
	
	//Produce an output array with filtered columns and rows (TAMs)
	public static ArrayList<ArrayList<String>> filter(ArrayList<ArrayList<String>> array) {
		ArrayList<Integer> columnNumbers = checkColumns(array);
		ArrayList<Integer> rowNumbers = checkRows(array);
		ArrayList<ArrayList<String>> output = new ArrayList<ArrayList<String>>();
		
		int numberOfRows = array.size();
		int numberOfColumns = array.get(0).size();
		
		for(int i = 0; i < numberOfRows; i++) {
			//Check for row numbers
			if(rowNumbers.contains(i) || i == 0) {
				ArrayList<String> tempRow = new ArrayList<String>();
			
				for(int j = 0; j < numberOfColumns; j++) {
					if(columnNumbers.contains(j)) {
						tempRow.add(array.get(i).get(j));
					}
				}
				
				output.add(tempRow);
			}
		}
		
		return output;
	}
	
	//Filter the first 11 columns
	public static ArrayList<ArrayList<String>> filter2(ArrayList<ArrayList<String>> array) {
		ArrayList<ArrayList<String>> output = new ArrayList<ArrayList<String>>();
		
		int rowNumbers = array.size();
		
		for(int i = 0; i < rowNumbers; i++) {
			ArrayList<String> temp = new ArrayList<String>();
			for(int j = 0; j < columnsNeeded.length; j++) {
				temp.add(array.get(i).get(j));
			}
			output.add(temp);
		}
		
		return output;
	}
	
	//Filter the first 11 columns for cell style array;
	public static ArrayList<ArrayList<XSSFCellStyle>> filterStyle(ArrayList<ArrayList<XSSFCellStyle>> array) {
		ArrayList<ArrayList<XSSFCellStyle>> output = new ArrayList<ArrayList<XSSFCellStyle>>();
		
		int rowNumbers = array.size();
		
		for(int i = 0; i < rowNumbers; i++) {
			ArrayList<XSSFCellStyle> temp = new ArrayList<XSSFCellStyle>();
			for(int j = 0; j < columnsNeeded.length; j++) {
				temp.add(array.get(i).get(j));
			}
			output.add(temp);
		}
		return output;
	}
	
	//Sort/swap the columns
	public static void sortColumns(ArrayList<ArrayList<String>> array) {
		int columnSize = array.get(0).size();
		int rowSize = array.size();
		
		for(int i = 0; i < columnSize - 1; i++) {	
			if(!(array.get(0).get(i).equals(columnsNeeded[i]))) {	
				//look for index number of searched column/string
				for(int j = i + 1; j < columnSize; j++) {
					if(array.get(0).get(j).equals(columnsNeeded[i])) {
						for(int m = 0; m < rowSize; m ++) {
							Collections.swap(array.get(m), i, j);
							
						}
					}
				}
			}
		}
	}
	
	//Rename the columns
	public static void renameColumns(ArrayList<ArrayList<String>> array) {
		int columnSize = array.get(0).size();
		
		for(int i = 1; i < columnSize; i++) {
			if(!array.get(0).get(i).equals(miaColumns[i])) {
				array.get(0).set(i, miaColumns[i]);
			}
		}
	}
	
	public static String formatDate(String date) {
		int dateLength = date.length();
		String temp = date.substring(0, dateLength - 2) + 
				"20" + date.substring(dateLength - 2, dateLength);
		
		return temp;
	}
	
	//Change the date format from 23/06/17 to 23/06/2017
	public static void formatDate_(ArrayList<String> array, int columnSize) {
		String temp1 = formatDate(array.get(columnSize));
		String temp2 = formatDate(array.get(columnSize - 1));
		array.set(columnSize, temp1);
		array.set(columnSize - 1, temp2);
	}
	
	//Correct the price format
	public static void correctPrice(ArrayList<String> array) {
		String temp = array.get(6);
		int length = temp.length();
		array.set(6, temp.substring(0, length - 3));
	}
	
	//Convert all TAM names to lower case
	public static void convertLowerCase(ArrayList<String> array) {
		String temp = array.get(1);
		array.set(1, temp.toLowerCase());
	}
	
	//Compare rows, if they are equal, returns true. False otherwise
	public static boolean compareRows(ArrayList<String> array1, ArrayList<String> array2) {
		int columnSize = array2.size();
		
		boolean b1 = true;
		for(int i = 0; i < columnSize; i++) {
			if(!array1.get(i).equals(array2.get(i))) {
				b1 = false;
			}
		}
		
		return b1;
	}
	
	//Check if the row is in the excel sheet array1, returns true if it's there. false otherwise
	public static boolean checkExistence(ArrayList<ArrayList<String>> array1, ArrayList<String> array2) {
		int rowSize = array1.size();
		
		boolean b1 = false;
		
		for(int i = 0; i < rowSize; i++) {
			if(compareRows(array1.get(i), array2)) {
				b1 = true;
			}
		}
		
		return b1;
	}
	
	//Add new rows from Unicorn to Mia's excel sheet. Keep the expired ones in Mia's
	public static void addRows(ArrayList<ArrayList<String>> array1, ArrayList<ArrayList<String>> array2) {
		int rowSize = array1.size();
		
		for(int i = 0; i < rowSize; i++) {
			if(!checkExistence(array2, array1.get(i))) {
				array2.add(array1.get(i));
			}
		}
	}
	
	//Error code 1 = there are more/less than 2 excel sheets
	//Create a text file with error message
	public static void printError(String errorMessage) {
		String timeStamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(Calendar.getInstance().getTime());
		
		String year = timeStamp.substring(0, 4);
		String month = timeStamp.substring(4, 6);
		String day = timeStamp.substring(6, 8);
		String date = "Date: " + day + "/" + month + "/" + year;
		
		String hour = timeStamp.substring(9, 11);
		String minute = timeStamp.substring(11, 13);
		String second = timeStamp.substring(timeStamp.length()-2);
		String time = "Time: " + hour + ":" + minute + ":" + second;
		
		//System.out.println(date);
		//System.out.println(time);
	}
	
	//Print excel sheets
	public static void printArray(ArrayList<ArrayList<String>> array) {
		int rows = array.size();
		int columns = array.get(0).size();
		
		for(int i = 0; i < rows; i++) {
			for(int j = 0; j < columns; j++) {
				System.out.print(array.get(i).get(j) + " | ");
			}
			System.out.println();
		}
	}
	
	//List all the files and folders in the directory. Won't traverse through subfolders
	public static void listFilesForFolder(final File folder) {
	    for(final File fileEntry : folder.listFiles()) {
	            System.out.println(fileEntry.getName());
	    }
	}
	
	//Get all excel sheets in the working directory. Add them to an array list
	public static ArrayList<String> getAllExcel() {
		ArrayList<String> array = new ArrayList<String>();
		String currentFolder = System.getProperty("user.dir").replace("\\", "/");
		final File folder = new File(currentFolder);
		
		//Loop through the folder. Won't traverse through subfolders. Put the names into the array
		//Will filter only excel files (.xlsx)
		for(final File fileEntry : folder.listFiles()) {
			//Will filter only excel files (.xlsx)
			String filename = fileEntry.getName();
			if(filename.length() >= 4) {
				if(filename.substring(filename.length()-4).equals("xlsx")
						&& !filename.substring(0, 2).equals("~$")) {
					String filepath = currentFolder + "/" + filename;
					array.add(filepath);
				}
			}
		}

		return array;
	}
	
	//Check if the two excel sheets are opened. True if it's opened.
	public static Boolean checkExcelSheetRunning(String filename) {
		Boolean b1 = false;
		
		File file = new File(filename);
		File sameFileName = new File(filename);
		
		if(!file.renameTo(sameFileName)) {
			//File is opened
			return true;
		}
		
		return b1;
	}
	
	//Check number of excel sheets. If there are 2, proceed. Abort if otherwise
	public static void checkExcel(ArrayList<String> array) {
		if(array.size() != 2) {
			//write log to file
			printError("1");
			System.exit(0);
		}
	}
	
	//get and set method
	public static void setTAM(String[] tam) {
		HTAMs = tam;
	}
	
	public static String[] getTam() {
		return HTAMs;
	}
	
	public static void setFile1(String filepath) {
		file1 = filepath;
	}
	
	public static String getFile1() {
		return file1;
	}
	
	public static void setFile2(String filepath) {
		file2 = filepath;
	}
	
	public static String getFile2() {
		return file2;
	}
	
	public static void setColumns(String[] columns) {
		columnsNeeded = columns;
	}
	
	public String[] getColumns() {
		return columnsNeeded;
	}
	
	public static void setOutputName(String name) {
		outputFile = name;
	}
	
	public String getOutputName() {
		return outputFile;
	}
	
	public static String convert(String[] s) {
		String output = s[0];
		
		for(int i = 1; i < s.length; i++) {
			output = output + ", " + s[i];
		}
		
		return output;
	}
}
