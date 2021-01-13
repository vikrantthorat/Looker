/**
 * This code will compare the values from Google file and Looker file to generate the output for mismatch records.
 */

package ClCompare.comparision_xcl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {
	static Boolean check = false;

	//Change column number whatever you want to take data
	public static void main(String args[]) throws IOException, ParseException {

		//Below List of array will capture the mismatch records and put it in output excel sheet.
		List<String> outputArray = new ArrayList<String>();
		outputArray.add("*************Mismatched records******************");

		//Below List of array will capture the mismatch primary keys only and put it in output excel sheet.
		List<String> outputArrayForGoogleKeyOnly = new ArrayList<String>();
		outputArrayForGoogleKeyOnly.add("*************Key present in google file but not in Looker file******************");
		
		//Below List of array will capture the mismatch primary keys only and put it in output excel sheet.
		List<String> outputArrayForlookerKeyOnly = new ArrayList<String>();
		outputArrayForlookerKeyOnly.add("*************Key present in Looker file but not in Google file******************");
				
		String file1 ="C:\\Users\\Dell\\Desktop\\Docs\\Looker\\Google - 26K.xlsx";
		String file2 = "C:\\Users\\Dell\\Desktop\\Docs\\Looker\\Looker - 26K.xlsx";

		try {

			Map<String, String> mpGoogle = findColumn(file1);
			Map<String, String> mapLooker = findColumn(file2);

			Set<String> keySet1 = mpGoogle.keySet();
			Set<String> keySet2 = mapLooker.keySet();

			//Convert set to array for better traversing.
			Object[] keyArr1 =keySet1.toArray();
			Object[] keyArr2 =keySet2.toArray();

			int matchedRecordCount =0;
			int misMatchRecordCount =0;
			int keyNotFound =0;
			
			//Below code will verify if individual record from Google sheet is avaialable in looker sheet and it is aviable then other columns are matching are not.
			if(keyArr1.length > keyArr2.length || keyArr1.length == keyArr2.length) {
				for (int m=0;m<keyArr1.length;m++)
				{
					for (int n=0;n<keyArr2.length;n++)
					{
						if(keyArr1[m].equals(keyArr2[n]))
						{
							if(mpGoogle.get(keyArr1[m]).equals(mapLooker.get(keyArr2[n])))
							{
								matchedRecordCount++;
							}
							else {
								outputArray.add("For Key -"+(String) (keyArr1[m])+" Data in Google Sheet data is - "+mpGoogle.get(keyArr1[m]).toString()+" And in Looker Sheet data is - "+mapLooker.get(keyArr2[n]).toString());
							}}else {
								misMatchRecordCount=(mpGoogle.size()-(matchedRecordCount+1));
							}
					}
				}
				System.out.println("Total Number of records in Google sheet - "+keyArr1.length);
				System.out.println("Total Number of records in Looker sheet - "+keyArr2.length);
			}else if(keyArr1.length < keyArr2.length) {
				for (int m=0;m<keyArr2.length;m++)
				{
					for (int n=0;n<keyArr1.length;n++)
					{
						if(keyArr2[m].equals(keyArr1[n]))
						{
							if(mapLooker.get(keyArr2[m]).equals(mpGoogle.get(keyArr1[n])))
							{
								matchedRecordCount++;
							}
							else {
								outputArray.add("For Key -"+(String) (keyArr2[m])+" Data in Google Sheet data is - "+mpGoogle.get(keyArr1[n]).toString()+" And in Looker Sheet data is - "+mapLooker.get(keyArr2[m]).toString());
							}
						}else {
							misMatchRecordCount=(mpGoogle.size()-(matchedRecordCount+1));
						}
					}
				}
				System.out.println("Total Number of records in Google sheet - "+keyArr1.length);
				System.out.println("Total Number of records in Looker sheet - "+keyArr2.length);
			}else {
				System.out.println("Total Number of records in Google sheet - "+keyArr1.length);
				System.out.println("Total Number of records in Looker sheet - "+keyArr2.length);
			}
			System.out.println("Total number of mismatched records -" + misMatchRecordCount);
			System.out.println("Total number of matched records -" + matchedRecordCount);

			//Below loop will verify if any key is not present in looker sheet.
			for(int k =0;k<keyArr1.length;k++)
			{
				if(!keySet2.contains(keyArr1[k]))
				{
					outputArrayForGoogleKeyOnly.add("Key - "+(String) (keyArr1[k])+" is not present in Looker Sheet.");
				}
			}
			
			//Below loop will verify if any key is not present in Google sheet.
			for(int k =0;k<keyArr2.length;k++)
			{
				if(!keySet1.contains(keyArr2[k]))
				{
					if(!outputArrayForGoogleKeyOnly.contains("Key - "+(String) (keyArr2[k])+" is not present in Looker Sheet."))
					{
						outputArrayForlookerKeyOnly.add("Key - "+(String) (keyArr2[k])+" is not present in Google Sheet.");
					}
				}
			}
			
			//This will write data into excel. 
			WriteDataToExcel(outputArray, outputArrayForGoogleKeyOnly,outputArrayForlookerKeyOnly);

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static Map<String, String> findColumn(String filePath) throws IOException, ParseException
	{
		Map<String, String> myMap1 = new HashMap<>();
		FileInputStream file = new FileInputStream(new File(filePath));

		//Below arrays will capture the values from sheet for respective columns.
		ArrayList compColArr = new ArrayList();
		ArrayList totalImpressionArr = new ArrayList();
		ArrayList totalClicksArr = new ArrayList();
		ArrayList totalCostArr = new ArrayList();
		ArrayList dateArr = new ArrayList();

		//subMap will be used to capture the values of other than "Comparison column".
		Map<Object, String> subMap = new HashMap<>();

		// Get the workbook instance for XLSX file
		@SuppressWarnings("resource")
		XSSFWorkbook workbook1 = new XSSFWorkbook(file);

		// Get only first sheet from the workbook
		XSSFSheet sheet1 = workbook1.getSheetAt(0);

		//Below for loop will capture the values of each collumns[e.g. comparision column] and add the values into respective arrays[eg compColArr] 
		for (int i =0;i<sheet1.getRow(0).getPhysicalNumberOfCells();i++)
		{
			switch(sheet1.getRow(0).getCell(i).getStringCellValue())
			{
			case "Comparison column":
				compColArr = addValuesToArray(sheet1, i);
				break;
			case "AdWords AdGroup Total Impressions":
				totalImpressionArr = addValuesToArray(sheet1, i);
				break;
			case "AdWords AdGroup Total Clicks":
				totalClicksArr = addValuesToArray(sheet1, i);
				break;
			case "AdWords AdGroup Total Cost":
				totalCostArr = addValuesToArray(sheet1, i);
				break;
			case "Date" :
				dateArr = addDateValueToArray(sheet1, i);
			}
		}

		//Below code will put the values in map.
		for (int j =0;j<sheet1.getLastRowNum();j++) {
			// Below line of code will add valued of other than "Comparison column" in "subMap" map.
			subMap.put(totalImpressionArr.get(j), (totalImpressionArr.get(j).toString().concat(", "+ totalClicksArr.get(j).toString().concat(", "+ totalCostArr.get(j).toString()))));
			//Below code will add values in map [eg - comparison column as a key and other columns value as a values]
			myMap1.put( (String) compColArr.get(j).toString().concat(dateArr.get(j).toString()), subMap.get(totalImpressionArr.get(j)));
			subMap.clear();
		}
		return myMap1;
	}

	//Below method will add values to arraylist from excel sheet
	public static ArrayList<String> addValuesToArray(XSSFSheet sheetName, int k) {
		ArrayList<String> arrayListValues = new ArrayList<String>();
		int j =0;
		int i = k;
		{
			for(j=1;j<sheetName.getLastRowNum()+1;j++)
			{
				switch (sheetName.getRow(j).getCell(i).getCellType())
				{
				case NUMERIC:
					DataFormatter formatter = new DataFormatter();
					String val = formatter.formatCellValue(sheetName.getRow(j).getCell(i));
					arrayListValues.add(val);
					break;
				case STRING:
					arrayListValues.add(sheetName.getRow(j).getCell(i).getStringCellValue());
					break;
				case BOOLEAN:
					arrayListValues.add(sheetName.getRow(j).getCell(i).getStringCellValue());
					break;
				default:
					break;	

				}
			}
		}
		return arrayListValues;
	}


	//Below method will add values to arraylist from excel sheet
	public static ArrayList<String> addDateValueToArray(XSSFSheet sheet1, int k) throws ParseException {
		ArrayList<String> arrayDateListValue = new ArrayList<String>();
		int i = k;
		{
			for(int j=1;j<sheet1.getLastRowNum()+1;j++)
			{
				switch (sheet1.getRow(j).getCell(i).getCellType())
				{
				case STRING:
					DateFormat outputFormat = new SimpleDateFormat("ddMMM", Locale.ENGLISH);
					String inputText = sheet1.getRow(j).getCell(i).getStringCellValue();
					Date date=new SimpleDateFormat("dd/MM/yyyy").parse(inputText); 
					String outputText = outputFormat.format(date);
					arrayDateListValue.add(outputText);
					break;
				default:
					break;

				}
			}
		}
		return arrayDateListValue;
	}


	//Below code will write output to excel sheet.
	public static void  WriteDataToExcel(List<String> outputarr1, List<String> outputarr2, List<String> outputarr3) throws IOException {
		//Create blank workbook
		@SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook();

		//Create a blank sheets
		XSSFSheet spreadsheet = workbook.createSheet( "Employee Info ");

		//Create row object
		XSSFRow row;

		//Iterate over data and write to sheet for mismatch keys + records in first column
		int rowid = 0;
		for (int i =0; i<outputarr1.size(); i++) {
			int cellid = 0;
			row = spreadsheet.createRow(rowid++);
			Cell cell = row.createCell(cellid++);
			cell.setCellValue((String) outputarr1.get(i));
		}
		
		//Iterate over data and write to sheet for mismatch Google keys only in same column
		int rowid2 = outputarr1.size();
		for (int i =0; i<outputarr2.size(); i++) {
			int cellid = 0;
			row = spreadsheet.createRow(rowid2++);
			Cell cell = row.createCell(cellid++);
			cell.setCellValue((String) outputarr2.get(i));
		}
	
		//Iterate over data and write to sheet for mismatch Google keys only in same column
				int rowid3 = outputarr1.size()+outputarr2.size();
				for (int i =0; i<outputarr3.size(); i++) {
					int cellid = 0;
					row = spreadsheet.createRow(rowid3++);
					Cell cell = row.createCell(cellid++);
					cell.setCellValue((String) outputarr3.get(i));
				}
				
		//Write the workbook in file system
		FileOutputStream out = new FileOutputStream(
				new File("C:\\Users\\Dell\\Desktop\\Docs\\Looker\\Writesheet1.xlsx"));

		workbook.write(out);
		out.close();
		System.out.println("Writesheet.xlsx written successfully");
	}

}
