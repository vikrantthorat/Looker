package ClCompare.comparision_xcl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hpsf.Array;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App {
		static Boolean check = false;
		
		//Change column number whatever you want to take data
		public static int columnNumForFirst = 0;
		public static int columnNumForSecond = 0;
		public static int columnNum2ForFirst = 1;
		public static int columnNum2ForSecond = 1;

		public static void main(String args[]) throws IOException, ParseException {

			try {
				//ArrayList arr3 = new ArrayList();
				List<String> arr3 = new ArrayList<String>();
				ArrayList<String> arr4 = new ArrayList<String>(50000);

				String file1 ="C:\\Users\\Dell\\Desktop\\Docs\\Looker\\Google - 26K.xlsx";

				String file2 = "C:\\Users\\Dell\\Desktop\\Docs\\Looker\\Looker - 26K.xlsx";

				Map<String, String> mapArr1 = findColumn(file1);
				Map<String, String> mapArr2 = findColumn(file2);
				//Map<String, List> arr1 = StoreArraysToHashMap(file1,columnNumForFirst);
				//Map<String, List> arr2 = StoreArraysToHashMap(file2,columnNumForFirst);

				
					Set<String> keysarr1 = mapArr1.keySet();
					Set<String> keysarr2 = mapArr2.keySet();
					Collection<String> valuesarr1 = mapArr1.values();
					Collection<String> valuesarr2 = mapArr2.values();
					Iterator i = keysarr1.iterator();
					Object[] ka1 =keysarr1.toArray();
					Object[] ka2 =keysarr2.toArray();
					int k =0;
					int l =0;
					if(ka1.length == ka2.length) {
					for (int m=0;m<ka1.length;m++)
					{
						for (int n=0;n<ka2.length;n++)
						{
							if(ka1[m].equals(ka2[n]))
							{
								if(mapArr1.get(ka1[m]).equals(mapArr2.get(ka2[n])))
								{
									k++;
								}
							else {
								arr3.add("For Key -"+(String) (ka1[m])+" Data in Google Sheet data is - "+mapArr1.get(ka1[m]).toString()+" And in Looker Sheet data is - "+mapArr2.get(ka2[n]).toString());
							}}else {
								l=(mapArr1.size()-(k+1));
							}
						}
					}
					}else {
						System.out.println("Total Number of records in Google sheet - "+ka1.length);
						System.out.println("Total Number of records in Looker sheet - "+ka2.length);
					}
				System.out.println("Total number of mismatched records -" + l);
				System.out.println("Total number of matched records -" + k);
				
				WriteDataToExcel(arr3);
				  
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
			ArrayList arr = new ArrayList();
			ArrayList arr1 = new ArrayList();
			ArrayList arr2 = new ArrayList();
			ArrayList arr3 = new ArrayList();
			ArrayList arr4 = new ArrayList();
			Map<Object, String> subMap = new HashMap<>();
			
			List subArr = new ArrayList<String>();
			// Get the workbook instance for XLSX file
			XSSFWorkbook workbook1 = new XSSFWorkbook(file);

			// Get only first sheet from the workbook
			XSSFSheet sheet1 = workbook1.getSheetAt(0);

			
			// Get iterator to all the rows in current sheet1ZZ
			Iterator<Row> rowIterator1 = sheet1.iterator();
			//System.out.println(sheet1.getRow(0).getPhysicalNumberOfCells());
			String a = null;
			String b =  null;
			for (int i =0;i<sheet1.getRow(0).getPhysicalNumberOfCells();i++)
			{
				//if(sheet1.getRow(0).getCell(i).getStringCellValue().equals("Comparison column") && sheet1.getRow(0).getCell(1).getStringCellValue().equals("AdWords AdGroup Total Impressions"))
				switch(sheet1.getRow(0).getCell(i).getStringCellValue())
				{
					case "Comparison column":
						arr = addValueToHashMap(sheet1, i);
						break;
					case "AdWords AdGroup Total Impressions":
						arr1 = addValueToHashMap(sheet1, i);
						break;
					case "AdWords AdGroup Total Clicks":
						arr2 = addValueToHashMap(sheet1, i);
						break;
					case "AdWords AdGroup Total Cost":
						arr3 = addValueToHashMap(sheet1, i);
						break;
					case "Date" :
						arr4 = addDateValueToHashMap(sheet1, i);
				}
			}
			

			for (int j =0;j<sheet1.getLastRowNum();j++) {
				//subArr= arr.subList(1, 4);
				//arr4 = {arr1.get(j),arr2.get(j), arr3.get(j)};
				subMap.put(arr1.get(j), (arr1.get(j).toString().concat(", "+ arr2.get(j).toString().concat(", "+ arr3.get(j).toString()))));
				myMap1.put( (String) arr.get(j).toString().concat(arr4.get(j).toString()), subMap.get(arr1.get(j)));
				subMap.clear();
			}
			return myMap1;
			
		}
		
		public static ArrayList addValueToHashMap(XSSFSheet sheetName, int k) {
			Map<String, List> myMap = new HashMap<>();
			ArrayList ar = new ArrayList<String>();
			int j =0;
			int i = k;
			{
				for(j=1;j<sheetName.getLastRowNum()+1;j++)
				{
					//System.out.println(j);
					switch (sheetName.getRow(j).getCell(i).getCellType())
					{
					case NUMERIC:
						DataFormatter formatter = new DataFormatter();
						String val = formatter.formatCellValue(sheetName.getRow(j).getCell(i));
						ar.add(val);
						break;
					case STRING:
						ar.add(sheetName.getRow(j).getCell(i).getStringCellValue());
						break;
					case BOOLEAN:
						ar.add(sheetName.getRow(j).getCell(i).getStringCellValue());
						break;	
						
					}
				}
			}
			return ar;
		}
		
		
		public static ArrayList addDateValueToHashMap(XSSFSheet sheet1, int k) throws ParseException {
			Map<String, List> myMap = new HashMap<>();
			ArrayList ar = new ArrayList<String>();
			int i = k;
			{
				for(int j=1;j<sheet1.getLastRowNum()+1;j++)
				{
					switch (sheet1.getRow(j).getCell(i).getCellType())
					{
					case STRING:
						DateFormat outputFormat = new SimpleDateFormat("ddMMM", Locale.ENGLISH);
						DateFormat inputFormat = new SimpleDateFormat("DD/MM/YYYY", Locale.ENGLISH);

						String inputText = sheet1.getRow(j).getCell(i).getStringCellValue();
						Date date=new SimpleDateFormat("dd/MM/yyyy").parse(inputText); 
						String outputText = outputFormat.format(date);
						ar.add(outputText);
						break;
						
					}
				}
			}
			return ar;
		}
		
		
		
		//Store data to HashMap
		private static  Map<String, List> StoreArraysToHashMap(String path, int columnNum) throws IOException {
			FileInputStream file = new FileInputStream(new File(path));
			ArrayList arr = new ArrayList();
			// Get the workbook instance for XLSX file
			XSSFWorkbook workbook1 = new XSSFWorkbook(file);

			// Get only first sheet from the workbook
			XSSFSheet sheet1 = workbook1.getSheetAt(0);

			
			// Get iterator to all the rows in current sheet1ZZ
			Iterator<Row> rowIterator1 = sheet1.iterator();
			
			Map<String, List> myMap = new HashMap<>();
			Row row = rowIterator1.next();
			// For each row, iterate through all the columns
			Iterator<Cell> cellIterator1 = row.cellIterator();
			Cell cell = cellIterator1.next();
			if(sheet1.getRow(0).getCell(0).getStringCellValue().equals("Comparison column") && sheet1.getRow(0).getCell(1).getStringCellValue().equals("AdWords AdGroup Total Impressions"))
			{
				for(int i =0; i<1;i++)
				{
					for(int j=1;j<sheet1.getLastRowNum();j++)
					{
						ArrayList ar = new ArrayList<String>();
						List subArr = new ArrayList<String>();
						switch (sheet1.getRow(j).getCell(i).getCellType())
						{
						case NUMERIC:
							ar.add(sheet1.getRow(j).getCell(i).getStringCellValue());
							break;
						case STRING:
							ar.add(sheet1.getRow(j).getCell(i).getStringCellValue());
							DataFormatter formatter = new DataFormatter();
							String val = formatter.formatCellValue(sheet1.getRow(j).getCell(++i));
							ar.add(val);
							i--;
							break;
						case BOOLEAN:
							ar.add(sheet1.getRow(j).getCell(i).getStringCellValue());
							break;	
						}
						
						ar.add("f");
						subArr= ar.subList(1, 2);
						myMap.put( ar.get(0).toString(), subArr);
						//System.out.println("@@@@"+myMap);
						
					}
				}
				//for(Map.Entry m : myMap.entrySet()){    
				 //   System.out.println(m.getKey()+" "+m.getValue());    
				//}
				//System.out.println("--"+myMap.size());
				file.close();
			}
			return myMap;
		}
		
		public static void  WriteDataToExcel(List<String> arr3) throws IOException {
			 //Create blank workbook
		      XSSFWorkbook workbook = new XSSFWorkbook();
		      
		      //Create a blank sheets
		      XSSFSheet spreadsheet = workbook.createSheet( "Employee Info ");

		      //Create row object
		      XSSFRow row;

		      //Iterate over data and write to sheet
		      int rowid = 0;
		      
		      for (int i =0; i<arr3.size(); i++) {
		    	  int cellid = 0;
		         row = spreadsheet.createRow(rowid++);
		            Cell cell = row.createCell(cellid++);
		            cell.setCellValue((String) arr3.get(i));
		      }
		      //Write the workbook in file system
		      FileOutputStream out = new FileOutputStream(
		         new File("C:\\Users\\Dell\\Desktop\\Docs\\Looker\\Writesheet1.xlsx"));
		      
		      workbook.write(out);
		      out.close();
		      System.out.println("Writesheet.xlsx written successfully");
		      //System.out.println("arr3 list values, here arr1 has some values which arr2 DOES NOT have : " + arr3);
			   }
		
}