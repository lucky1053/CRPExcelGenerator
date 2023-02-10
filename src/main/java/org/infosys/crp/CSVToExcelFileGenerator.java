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
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Properties;
import java.util.Set;

public class CSVToExcelFileGenerator {

    FileInputStream fis;
    Properties prop;
    String inputFilePath;
    String timeStamp;
    String countryCode;
    String outputFilePath;

    public CSVToExcelFileGenerator() {
        try {
            fis = new FileInputStream(System.getProperty("user.dir") + "\\src\\main\\resources\\ApplicationProperty.properties");
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
        prop = new Properties();
        try {
            prop.load(fis);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        inputFilePath = prop.getProperty("InputFilePath");
        outputFilePath = prop.getProperty("OutputFilePath");
        timeStamp = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss").format(LocalDateTime.now());
    }

    public void inputCSVToExcel(String inputFileName) {
        try {
            XSSFWorkbook workBook = new XSSFWorkbook();
            XSSFSheet sheet = workBook.createSheet("sheet1");
            String currentLine;
            int RowNum = 0;
            BufferedReader reader = new BufferedReader(new FileReader(inputFilePath + "\\" + inputFileName));
            while ((currentLine = reader.readLine()) != null) {
                String[] data = currentLine.split(",");
                XSSFRow currentRow = sheet.createRow(RowNum);
                for (int index = 0; index < data.length; index++) {
                    currentRow.createCell(index).setCellValue(data[index]);
                }
                RowNum++;
            }
            String outputExcelFileName = inputFileName.substring(0, inputFileName.indexOf("."));
            FileOutputStream fileOutputStream = new FileOutputStream(inputFilePath + "\\" + outputExcelFileName + "_" + timeStamp + ".xlsx");
            reader.close();
            workBook.write(fileOutputStream);
            workBook.close();
            fileOutputStream.close();
            System.out.println(inputFileName + " file to " + outputExcelFileName + "_" + timeStamp + ".xlsx file Conversion is completed.");
            System.out.println(outputExcelFileName + "_" + timeStamp + ".xlsx file contains total AGREEMENT KEY records : " + (RowNum - 7));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public Set<String> onlySourceUniqueRecordSetGeneration() throws IOException {
        File file = new File(inputFilePath + "\\Only_Source_" + timeStamp + ".xlsx");
        FileInputStream fis = new FileInputStream(file);
        XSSFWorkbook workBook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workBook.getSheetAt(0);
        int total_record = sheet.getLastRowNum();
        System.out.println("total records in source only excel file : " + (total_record - 7 + 1));
        int noOfCharRemove = Integer.parseInt(prop.getProperty("NoOfCharacterRemove"));
        System.out.println("Character to be removed from Agreemnet Key's : " + noOfCharRemove);
        Set<String> setOfUniqueAgreementKey = new HashSet<>();
        try {
            XSSFRow r = sheet.getRow(7);
            String countryCellValue = r.getCell(0).getStringCellValue();
            countryCode = countryCellValue.substring(0, 2);
            for (int i = 7; i <= total_record; i++) {
                XSSFRow row = sheet.getRow(i);
                String cellValue = row.getCell(0).getStringCellValue();
                String updatedValue = cellValue.substring(noOfCharRemove, cellValue.length() - 3);
                setOfUniqueAgreementKey.add(updatedValue);
            }
            workBook.close();
            fis.close();
        } catch (Exception ex) {
            System.out.println(ex.getMessage() + "Exception in source only file reading");
        }
        return setOfUniqueAgreementKey;
    }

    public void sourceOnlyUniqueRecordExcelFileGeneration(Set<String> setOfSourceOnlyUniqueRecords) {
        try {
            XSSFWorkbook workBook = new XSSFWorkbook();
            XSSFSheet sheet = workBook.createSheet("Missing Data");
            XSSFRow row = sheet.createRow(0);
            XSSFCell cell = row.createCell(0);
            cell.setCellValue("Missing Data");
            int n = setOfSourceOnlyUniqueRecords.size();
            String[] arr = new String[n];
            int index = 0;
            for (String x : setOfSourceOnlyUniqueRecords)
                arr[index++] = x;
            for (int i = 1; i <= setOfSourceOnlyUniqueRecords.size(); i++) {
                XSSFRow rowData = sheet.createRow(i);
                XSSFCell cellData = rowData.createCell(0);
                cellData.setCellValue(arr[i - 1]);
            }
            File filenew = new File(outputFilePath + "\\" + countryCode + "_Only_Source_Missing_" + timeStamp + ".xlsx");
            FileOutputStream out = new FileOutputStream(filenew);
            workBook.write(out);
            workBook.close();
            out.close();
            System.out.println("Source Only excel file with unique records is completed.");
            System.out.println("Source Only excel file contains unique total AGREEMENT KEY records : " + arr.length);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void sourceOnlyUniqueRecordsExcelToMissingCSVGeneration() {
        try {
            File inputFile = new File(outputFilePath + "\\" + countryCode + "_Only_Source_Missing_" + timeStamp + ".xlsx");
            File outputFile = new File(outputFilePath + "\\" + countryCode + "_Only_Source_Missing_" + timeStamp + ".csv");
            StringBuilder data = new StringBuilder();
            try {
                FileInputStream fis = new FileInputStream(inputFile);
                Workbook workbook;
                if (inputFile.getName().endsWith(".xlsx")) {
                    workbook = new XSSFWorkbook(fis);
                } else if (inputFile.getName().endsWith(".xls")) {
                    workbook = new HSSFWorkbook(fis);
                } else {
                    fis.close();
                    throw new Exception("File not supported!");
                }
                Sheet sheet = workbook.getSheetAt(0);
                for (Row row : sheet) {
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
                    data.append('\n');
                }
                fis.close();
                FileOutputStream fos = new FileOutputStream(outputFile);
                fos.write(data.toString().getBytes());
                workbook.close();
                fos.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
            System.out.println("Conversion of " + countryCode + "_Only_Source_Missing_" + timeStamp +
                    ".xlsx file to " + countryCode + "_Only_Source_Missing_" + timeStamp + ".csv file is completed.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public Set<String> matchCsvToMatchExcelGeneration() throws IOException {
        File file = new File(inputFilePath + "\\Match_" + timeStamp + ".xlsx");
        FileInputStream fis = new FileInputStream(file);
        XSSFWorkbook workBook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workBook.getSheetAt(0);
        int total_record = sheet.getLastRowNum();
        System.out.println("total records in match excel file : " + (total_record - 7 + 1));
        Set<String> setOfUniqueAgreementKey = new HashSet<>();
        try {
            for (int i = 7; i <= total_record; i++) {
                XSSFRow row = sheet.getRow(i);
                String cellValue = row.getCell(0).getStringCellValue();
                setOfUniqueAgreementKey.add(cellValue);
            }
            System.out.println("Unique records count in match excel file " + setOfUniqueAgreementKey.size());
            workBook.close();
            fis.close();
        } catch (Exception ex) {
            System.out.println(ex.getMessage() + "Exception in match only file reading");
        }
        return setOfUniqueAgreementKey;
    }


    public void matchUniqueRecordExcelFileGeneration(Set<String> matchUniqueRecordSet) {
        try {
            XSSFWorkbook workBook = new XSSFWorkbook();
            XSSFSheet sheet = workBook.createSheet("Matching Data");
            XSSFRow row = sheet.createRow(0);
            XSSFCell cell = row.createCell(0);
            cell.setCellValue("Matching Data");
            int n = matchUniqueRecordSet.size();
            String[] arr = new String[n];
            int index = 0;
            for (String x : matchUniqueRecordSet)
                arr[index++] = x;
            for (int i = 1; i <= matchUniqueRecordSet.size(); i++) {
                XSSFRow rowData = sheet.createRow(i);
                XSSFCell cellData = rowData.createCell(0);
                cellData.setCellValue(arr[i - 1]);
            }
            File filenew = new File(outputFilePath + "\\" + countryCode + "_Matching_" + timeStamp + ".xlsx");
            FileOutputStream out = new FileOutputStream(filenew);
            workBook.write(out);
            workBook.close();
            out.close();
            System.out.println("Matching excel file with unique records is completed.");
            System.out.println("Matching excel file contains unique total AGREEMENT KEY records : " + arr.length);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    public void matchUniqueRecordsExcelToMatchingCSVGeneration() {
        try {
            File inputFile = new File(outputFilePath + "\\" + countryCode + "_Matching_" + timeStamp + ".xlsx");
            File outputFile = new File(outputFilePath + "\\" + countryCode + "_Matching_" + timeStamp + ".csv");
            StringBuilder data = new StringBuilder();
            try {
                FileInputStream fis = new FileInputStream(inputFile);
                Workbook workbook;
                if (inputFile.getName().endsWith(".xlsx")) {
                    workbook = new XSSFWorkbook(fis);
                } else if (inputFile.getName().endsWith(".xls")) {
                    workbook = new HSSFWorkbook(fis);
                } else {
                    fis.close();
                    throw new Exception("File not supported!");
                }
                Sheet sheet = workbook.getSheetAt(0);
                for (Row row : sheet) {
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
                    data.append('\n');
                }
                fis.close();
                FileOutputStream fos = new FileOutputStream(outputFile);
                fos.write(data.toString().getBytes());
                workbook.close();
                fos.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
            System.out.println("Conversion of " + countryCode + "_Matching_" + timeStamp +
                    ".xlsx file to " + countryCode + "_Matching_" + timeStamp + ".csv file is completed.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void excelFilesCleanUp() {
        inputFolderExcelFilesCleanUp();
        outputFolderExcelFilesCleanUp();
    }

    private void outputFolderExcelFilesCleanUp() {
        String path = outputFilePath + "\\";
        File fObj = new File(path);
        if (fObj.exists() && fObj.isDirectory()) {
            File[] a = fObj.listFiles();
            assert a != null;
            this.deleteFiles(a, 0);
        }
        System.out.println("All excel files are deleted from Output Files Folder.");
    }

    private void inputFolderExcelFilesCleanUp() {
        String inputFilepath = inputFilePath + "\\";
        File finputObj = new File(inputFilepath);
        if (finputObj.exists() && finputObj.isDirectory()) {
            File[] a = finputObj.listFiles();
            assert a != null;
            this.deleteFiles(a, 0);
        }
        System.out.println("All excel files are deleted from Input Files Folder.");
    }

    private void deleteFiles(File[] a, int i) {
        if (i == a.length) {
            return;
        }
        if (a[i].isFile()) {
            if (a[i].getName().contains(".xlsx")) {
                File file = a[i];
                file.delete();
            }
        }
        deleteFiles(a, i + 1);
    }
}
