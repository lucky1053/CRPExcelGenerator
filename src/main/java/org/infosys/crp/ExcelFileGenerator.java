package org.infosys.crp;

import java.util.Set;

public class ExcelFileGenerator {
    public static void main(String[] args) {
        try {
            CSVToExcelFileGenerator csv = new CSVToExcelFileGenerator();

            //Source Only Records file generation steps
            csv.inputCSVToExcel("Only_Source.csv");

            Set<String> sourceOnlyUniqueRecordSet= csv.onlySourceUniqueRecordSetGeneration();
            System.out.println("Source Only Unique Records : "+sourceOnlyUniqueRecordSet);

            csv.sourceOnlyUniqueRecordExcelFileGeneration(sourceOnlyUniqueRecordSet);

            csv.sourceOnlyUniqueRecordsExcelToMissingCSVGeneration();

            //Match file generation steps
            csv.inputCSVToExcel("Match.csv");

            Set<String> matchUniqueRecordSet=csv.matchCsvToMatchExcelGeneration();
            System.out.println("Match Unique Records : "+matchUniqueRecordSet);

            csv.matchUniqueRecordExcelFileGeneration(matchUniqueRecordSet);

            csv.matchUniqueRecordsExcelToMatchingCSVGeneration();

            //excel files clean steps
            csv.excelFilesCleanUp();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


}

