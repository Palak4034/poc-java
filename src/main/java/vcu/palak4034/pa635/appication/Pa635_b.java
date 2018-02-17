/**
 * 
 */
package vcu.palak4034.pa635.appication;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.Reader;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;

/**
 * @author palak
 *
 */
public class Pa635_b {

	/**
	 * @param args
	 */
	public static synchronized void main(String[] args) {
		System.out.println("Hi, programs starts.");

		Utility_b.convertCsvToExcel("C:\\Users\\palak\\Desktop\\KDDM\\Assignment\\dataset_missing10.csv");
		System.out.println("CSV file reading and conversion to Excel Program  is Done...!!!");

		System.out.println("reading file...");
		List<String> output = Utility_b.readFeatureFromXlsFile(5, null);
		System.out.println("Reading operation complete");

		System.out.println(output);

		double meanValue = Calculator_b.calculateMeanImputation(output);
		System.out.println("Calculated Mean Imputation is : " + meanValue);

		Utility_b.writeOutputCsvReplaceOccuranceFromInputFile("output_19.csv", "?", String.valueOf(meanValue),
				"C:\\Users\\palak\\eclipse-workspace\\assignment-pa635\\InputDataFomat_7.xls");

		System.out.println("Done.!");
	}

}

class Calculator_b {
	public static double calculateMeanImputation(List<String> inputData) {
		double sum = 0.0;
		int nonEmptyFieldsCount = 0;
		for (String input : inputData) {
			if (!input.trim().equalsIgnoreCase("?")) {
				sum = sum + (Double.parseDouble(input));
				nonEmptyFieldsCount++;
			}
		}
		return sum / nonEmptyFieldsCount;
	}
}

class Utility_b {

	public static boolean convertCsvToExcel(String csvFileName) {
		try {
			Workbook wb = new HSSFWorkbook();
			CreationHelper helper = wb.getCreationHelper();
			Sheet sheet = wb.createSheet("sheet1");

			CSVReader reader = new CSVReader(new FileReader(csvFileName));
			String[] line;
			int rowCount = 0;
			while ((line = reader.readNext()) != null) {
				Row row = sheet.createRow((short) rowCount++);

				for (int i = 0; i < line.length; i++) {
					
					if(!line[i].equalsIgnoreCase("?")) {
						row.createCell(i).setCellValue(helper.createRichTextString(line[i]));
						System.out.println("?:"+helper.createRichTextString(line[i]));
					}
					row.createCell(i).setCellValue((line[i]));
				}
			}

			// Write the output to a file
			FileOutputStream fileOut = new FileOutputStream("InputDataFomat_7.xls");
			wb.write(fileOut);
			fileOut.close();

			reader.close();
			wb.close();
		} catch (Exception exception) {
			System.out.println("Exception happened");
			return false;
		}
		return true;
	}

	public static List<String> readFeatureFromXlsFile(int featureColumn, String xlsFileName) {
		List<String> readFeatureList = new ArrayList<String>();
		try {
			InputStream ExcelFileToRead = new FileInputStream(
					"C:\\Users\\palak\\eclipse-workspace\\assignment-pa635\\InputDataFomat_2.xls");
			HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);

			HSSFSheet sheet = wb.getSheetAt(0);
			HSSFRow row;
			HSSFCell cell;

			Iterator<Row> rows = sheet.rowIterator();
			int rowCount = 1;

			while (rows.hasNext()) {
				row = (HSSFRow) rows.next();
				if (rowCount > 1) {
					Iterator<Cell> cells = row.cellIterator();

					for (int cellCount = 1; cells.hasNext(); cellCount++) {
						cell = (HSSFCell) cells.next();

						if (cellCount == featureColumn) {
							if (cell.getCellTypeEnum() == CellType.STRING) {
								readFeatureList.add(cell.getStringCellValue());
							} else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
								Double cellDoubleValue = cell.getNumericCellValue();
								readFeatureList.add(cellDoubleValue.toString());
							}
						}
					}
				}
				rowCount++;
			}
			wb.close();
		} catch (Exception exception) {
			System.out.println("Exception happened in reading the input xls file.");
			return null;
		}
		return readFeatureList;
	}

	public static boolean writeOutputCsvReplaceOccuranceFromInputFile(String csvFileName, String toBeReplaced,
			String replacingString, String excelFileName) {
		System.out.println("Writing output to an CSV File.");
		try {
			InputStream ExcelFileToRead = new FileInputStream(excelFileName);
			HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);

			HSSFSheet sheet = wb.getSheetAt(0);
			HSSFRow row;
			HSSFCell cell;

			Iterator<Row> rows = sheet.rowIterator();
			int rowCount = 1;

			CSVWriter csvWriter = new CSVWriter(new FileWriter(csvFileName));
			while (rows.hasNext()) {
				row = (HSSFRow) rows.next();
				if (rowCount > 1) {
					Iterator<Cell> cells = row.cellIterator();

					int columnCount = 1;
					String[] csvRow = new String[14];

					while (cells.hasNext()) {
						cell = (HSSFCell) cells.next();
						if (cell.getCellTypeEnum() == CellType.STRING) {
							if (cell.getStringCellValue().contains("/?".trim())) {
								csvRow[columnCount - 1] = replacingString;
							}
							csvRow[columnCount - 1] = cell.getStringCellValue();
						} else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
							Double cellDoubleValue = cell.getNumericCellValue();
							csvRow[columnCount - 1] = cellDoubleValue.toString();
						}
						columnCount++;
					}
					csvWriter.writeNext(csvRow);
				}
				rowCount++;
			}
			csvWriter.close();
			wb.close();
		} catch (Exception exception) {
			System.out.println("Exception : " + exception + " in writing values to File : " + csvFileName);
		}
		return true;
	}

}
