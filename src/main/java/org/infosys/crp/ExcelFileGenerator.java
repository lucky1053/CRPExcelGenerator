package org.infosys.crp;

import java.io.File;
import java.io.IOException;
import java.util.Set;

public class ExcelFileGenerator {
    public static void main(String[] args) throws IOException {
        try {
            String csvFileAddress = "Only_Source.csv";
            CSVToExcelConversion con = new CSVToExcelConversion();
            con.csvToExcel(csvFileAddress);
            Set<String> uniqueRecordSet = con.sourceXlsxToOutputXlsxConversion();
            con.uniqueRecordFileGenerator(uniqueRecordSet);
            con.uniqueRecordExcelToCsvConversion();
         //   FileDeletion fd = new FileDeletion();
       //     fd.deleteExcelFiles();

            String path = "D:\\Tutorial_Automation\\Java\\CRPExcelReportGenerator\\Output Files";
            File fObj = new File(path);
            ExcelFileGenerator obj = new ExcelFileGenerator();
            if (fObj.exists() && fObj.isDirectory()) {
                File a[] = fObj.listFiles();
                System.out.println("= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =");
                System.out.println("Displaying Files from the directory : " + fObj);
                System.out.println("= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =");
                obj.printFileNames(a, 0);
            }

            String path2 = "D:\\Tutorial_Automation\\Java\\CRPExcelReportGenerator\\Input Files";
            File fObj2 = new File(path2);
            ExcelFileGenerator obj2 = new ExcelFileGenerator();
            if (fObj2.exists() && fObj2.isDirectory()) {
                File a[] = fObj2.listFiles();
                System.out.println("= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =");
                System.out.println("Displaying Files from the directory : " + fObj2);
                System.out.println("= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =");
                obj2.printFileNames(a, 0);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void printFileNames(File[] a, int i) {
        if (i == a.length) {
            return;
        }
        if (a[i].isFile()) {
            System.out.println(a[i].getName());
            if(a[i].getName().contains(".xlsx")){
                a[i].delete();
            }
        }
        printFileNames(a, i + 1);
    }
}

