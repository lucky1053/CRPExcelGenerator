package org.infosys.crp;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;

public class CSVToExcelConversion {
    XSSFWorkbook workBook;
    XSSFSheet sheet;
    FileOutputStream fileOutputStream;
    static String cuurentDirPath;
    FileInputStream fis;
    XSSFRow row;
    XSSFCell cell;
    Set<String> setOfUniqueAgreementKey;

    public CSVToExcelConversion() {
        this.cuurentDirPath = System.getProperty("user.dir");
    }

    public void csvToExcel(String csvFileAddress) {
        try {
            workBook = new XSSFWorkbook();
            sheet = workBook.createSheet("sheet1");
            String currentLine;
            int RowNum = 0;
            BufferedReader br = new BufferedReader(new FileReader(cuurentDirPath + "\\Input Files\\" + csvFileAddress));
            while ((currentLine = br.readLine()) != null) {
                String[] str = currentLine.split(",");
                XSSFRow currentRow = sheet.createRow(RowNum);
                for (int i = 0; i < str.length; i++) {
                    currentRow.createCell(i).setCellValue(str[i]);
                }
                RowNum++;
            }
            fileOutputStream = new FileOutputStream(cuurentDirPath + "\\Input Files\\Only_Source.xlsx");
            workBook.write(fileOutputStream);
            fileOutputStream.close();
            System.out.println("Source CSV file to EXCEL file Conversion is completed.");
            System.out.println("Excel file contains total AGREEMENT KEY records : " + (RowNum - 7));
        } catch (Exception ex) {
            System.out.println(ex.getMessage() + "Exception in try");
        }
    }

    public Set<String> sourceXlsxToOutputXlsxConversion() throws IOException {
        File file = new File(cuurentDirPath + "\\Input Files\\Only_Source.xlsx");
        fis = new FileInputStream(file);
        workBook = new XSSFWorkbook(fis);
        sheet = workBook.getSheetAt(0);
        int total_record = sheet.getLastRowNum();
        System.out.println("total_record : " + total_record);
        try {
            setOfUniqueAgreementKey = new HashSet<>();
            for (int i = 7; i <= total_record; i++) {
                XSSFRow row = sheet.getRow(i);
                String cellValue = row.getCell(0).getStringCellValue();
                String updatedValue = cellValue.substring(5, cellValue.length() - 3);
                setOfUniqueAgreementKey.add(updatedValue);
            }
        } catch (Exception ex) {
            System.out.println(ex.getMessage() + "Exception in file reading");
        }
        return setOfUniqueAgreementKey;
    }

    public void uniqueRecordFileGenerator(Set<String> setOfUniqueAgreementKey) {
        try {
            workBook = new XSSFWorkbook();
            sheet = workBook.createSheet("Missing Data");
            row = sheet.createRow(0);
            cell = sheet.getRow(0).createCell(0);
            cell.setCellValue("Missing Data");
            int n = setOfUniqueAgreementKey.size();
            String[] arr = new String[n];
            int index = 0;
            for (String x : setOfUniqueAgreementKey)
                arr[index++] = x;
            for (int i = 1; i <= setOfUniqueAgreementKey.size(); i++) {
                XSSFRow rowData = sheet.createRow(i);
                XSSFCell cellData = rowData.createCell(0);
                cellData.setCellValue(arr[i - 1]);
            }
            File filenew = new File(cuurentDirPath + "\\Output Files\\Only_Source_Unique.xlsx");
            FileOutputStream out = new FileOutputStream(filenew);
            workBook.write(out);
            out.close();
            System.out.println("Source Only excel file with unique records is completed.");
            System.out.println("Source Only Excel file contains unique total AGREEMENT KEY records : " + arr.length);
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    public void uniqueRecordExcelToCsvConversion() {
        try {
            File inputFile = new File(cuurentDirPath + "\\Output Files\\Only_Source_Unique.xlsx");
            // Creating a outputFile object to write excel data to csv
            File outputFile = new File(cuurentDirPath + "\\Output Files\\Only_Source_Unique.csv");
            // For storing data into CSV files
            StringBuilder data = new StringBuilder();
            try {
                // Creating input stream
                FileInputStream fis = new FileInputStream(inputFile);
                Workbook workbook;
                // Get the workbook object for Excel file based on file format
                if (inputFile.getName().endsWith(".xlsx")) {
                    workbook = new XSSFWorkbook(fis);
                } else if (inputFile.getName().endsWith(".xls")) {
                    workbook = new HSSFWorkbook(fis);
                } else {
                    fis.close();
                    throw new Exception("File not supported!");
                }
                // Get first sheet from the workbook
                Sheet sheet = workbook.getSheetAt(0);
                // Iterate through each rows from first sheet
                for (Row row : sheet) {
                    // For each row, iterate through each columns
                    Iterator<Cell> cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        switch (cell.getCellType()) {
                            case BOOLEAN -> data.append(cell.getBooleanCellValue()).append(",");
                            case NUMERIC -> data.append(cell.getNumericCellValue()).append(",");
                            case STRING -> data.append(cell.getStringCellValue()).append(",");
                            case BLANK -> data.append("" + ",");
                            default -> data.append(cell).append(",");
                        }
                    }
                    // appending new line after each row
                    data.append('\n');
                }
                FileOutputStream fos = new FileOutputStream(outputFile);
                fos.write(data.toString().getBytes());
                fos.close();

            } catch (Exception e) {
                e.printStackTrace();
            }
            System.out.println("Conversion of an Excel file to CSV file is done!");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


}
