package ClCompare.comparision_xcl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.hpsf.Array;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App2 {
		static Boolean check = false;
		
		//Change column number whatever you want to take data
		public static int columnNumForFirst = 0;
		public static int columnNumForSecond = 0;
		public static int columnNum2ForFirst = 1;
		public static int columnNum2ForSecond = 1;

		public static void main(String args[]) throws IOException {

			try {

				ArrayList arr1 = new ArrayList();
				ArrayList arr2 = new ArrayList();
				//ArrayList arr3 = new ArrayList();
				ArrayList<String> arr3 = new ArrayList<String>();
				ArrayList<String> arr4 = new ArrayList<String>();

				FileInputStream file1 = new FileInputStream(new File(
						"C:\\Users\\Dell\\Desktop\\Docs\\Looker\\Google - 26K.xlsx"));

				FileInputStream file2 = new FileInputStream(new File(
						"C:\\Users\\Dell\\Desktop\\Docs\\Looker\\Looker - 26K.xlsx"));

				// Get the workbook instance for XLSX file
				XSSFWorkbook workbook1 = new XSSFWorkbook(file1);
				XSSFWorkbook workbook2 = new XSSFWorkbook(file2);

				// Get only first sheet from the workbook
				XSSFSheet sheet1 = workbook1.getSheetAt(0);
				XSSFSheet sheet2 = workbook2.getSheetAt(0);

				
				// Get iterator to all the rows in current sheet1
				Iterator<Row> rowIterator1 = sheet1.iterator();
				Iterator<Row> rowIterator2 = sheet2.iterator();
				
				//getting date from first excel file
				while (rowIterator1.hasNext()) {
					Row row = rowIterator1.next();
					// For each row, iterate through all the columns
					Iterator<Cell> cellIterator = row.cellIterator();

					while (cellIterator.hasNext()) {

						Cell cell = cellIterator.next();

						// This is for read only one column from excel
						if (cell.getColumnIndex() == columnNumForFirst) {
							// Check the cell type and format accordingly
							switch (cell.getCellType()) {
							case NUMERIC:
								//System.out.print(cell.getNumericCellValue());
								arr1.add(cell.getNumericCellValue());
								break;
							case STRING:
								arr1.add(cell.getStringCellValue());
								//System.out.print(cell.getStringCellValue());
								break;
							case BOOLEAN:
								arr1.add(cell.getBooleanCellValue());
								//System.out.print(cell.getBooleanCellValue());
								break;
							}

						}

					}

					//System.out.println(" ");
				}

				file1.close();

				//System.out.println("\n-----------------------------------");
				// For retrive the second excel data
				while (rowIterator2.hasNext()) {
					Row row1 = rowIterator2.next();
					// For each row, iterate through all the columns
					Iterator<Cell> cellIterator1 = row1.cellIterator();

					while (cellIterator1.hasNext()) {

						Cell cell1 = cellIterator1.next();
						// Check the cell type and format accordingly

						// This is for read only one column from excel
						if (cell1.getColumnIndex() == columnNumForSecond) {
							switch (cell1.getCellType()) {
							case NUMERIC:
								arr2.add(cell1.getNumericCellValue());
								//System.out.print(cell1.getNumericCellValue());
								break;
							case STRING:
								arr2.add(cell1.getStringCellValue());
								//System.out.print(cell1.getStringCellValue());
								break;
							case BOOLEAN:
								arr2.add(cell1.getBooleanCellValue());
								//System.out.print(cell1.getBooleanCellValue());
								break;

							}

						}
						// this continue is for
						// continue;
					}

					//System.out.println("");
				}

				System.out.println("Total Number of records in Google.xlsx -- " + arr1.size());
				System.out.println("Total Number of records in Looker.xlsx -- " + arr2.size());

				// compare two arrays
				for (Object process : arr1) {
					if (!arr2.contains(process)) {
						arr3.add(process.toString());
					}else {
						arr4.add(process.toString());
					}
				}
				//StoreArraysToHashMap(arr1, arr2);	

				// closing the files
				file1.close();
				file2.close();
				
				   //Create blank workbook
			      XSSFWorkbook workbook = new XSSFWorkbook();
			      
			      //Create a blank sheet
			      XSSFSheet spreadsheet = workbook.createSheet( "Employee Info ");

			      //Create row object
			      XSSFRow row;

			      //Iterate over data and write to sheet
			      int rowid = 0;
			      
			      for (int i =0; i<arr3.size(); i++) {
			    	  int cellid = 0;
			         row = spreadsheet.createRow(rowid++);
			            Cell cell = row.createCell(cellid++);
			            cell.setCellValue(arr3.get(i));
			      }
			      //Write the workbook in file system
			      FileOutputStream out = new FileOutputStream(
			         new File("C:\\Users\\Dell\\Desktop\\Docs\\Looker\\Writesheet1.xlsx"));
			      
			      workbook.write(out);
			      out.close();
			      System.out.println("Writesheet.xlsx written successfully");
			      //System.out.println("arr3 list values, here arr1 has some values which arr2 DOES NOT have : " + arr3);
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
			

		}

		//Store data to HashMap
		private  void StoreArraysToHashMap(ArrayList arr1, ArrayList arr2) {

			HashMap<Integer, String> hashMap1 = new HashMap<Integer, String>();
			for(int i = 0; i < arr1.size(); i++) {
				hashMap1.put(i, arr1.get(i).toString());
			}

			HashMap<Integer, String> hashMap2 = new HashMap<Integer, String>();
			for(int j = 0; j < arr1.size(); j++) {
				hashMap2.put(j, arr2.get(j).toString());
			}
			
			System.out.println("\nHashMap from first excel file: " + hashMap1);
			System.out.println("\nHashMap from second excel file: " + hashMap2);
			
		}
		
		public static void  WriteDataToExcel() throws IOException {

			      //Create blank workbook
			      XSSFWorkbook workbook = new XSSFWorkbook();
			      
			      //Create a blank sheet
			      XSSFSheet spreadsheet = workbook.createSheet( "Employee Info ");

			      //Create row object
			      XSSFRow row;
			      ArrayList<String> arr3 = new ArrayList<String>();
			      //Iterate over data and write to sheet
			      int rowid = 0;
			      
			      for (int i =0; i<arr3.size(); i++) {
			         row = spreadsheet.createRow(rowid++);
			         int cellid = 0;
			            Cell cell = row.createCell(cellid++);
			            cell.setCellValue(arr3.get(0));
			      }
			      //Write the workbook in file system
			      FileOutputStream out = new FileOutputStream(
			         new File("C:/poiexcel/Writesheet.xlsx"));
			      
			      workbook.write(out);
			      out.close();
			      System.out.println("Writesheet.xlsx written successfully");
			   }

			}
