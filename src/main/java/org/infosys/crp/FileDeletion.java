package org.infosys.crp;

import java.io.File;

public class FileDeletion {

    public void deleteExcelFiles() {
        String path = "D:\\Tutorial_Automation\\Java\\CRPExcelReportGenerator\\Output Files\\";
        File fObj = new File(path);

        if (fObj.exists() && fObj.isDirectory()) {
            File a[] = fObj.listFiles();
            System.out.println("= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =");
            System.out.println("Displaying Files from the directory : " + fObj);
            System.out.println("= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =");
            this.printFileNames(a, 0);
        }

        String inputFilepath = "D:\\Tutorial_Automation\\Java\\CRPExcelReportGenerator\\Input Files";
        File finputObj = new File(inputFilepath);


        if (finputObj.exists() && finputObj.isDirectory()) {
            File a[] = finputObj.listFiles();
            System.out.println("= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =");
            System.out.println("Displaying Files from the directory : " + finputObj);
            System.out.println("= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =");
            this.printFileNames(a, 0);
        }
    }

    public void printFileNames(File[] a, int i) {
        if (i == a.length) {
            return;
        }
        if (a[i].isFile()) {
           // System.out.println(a[i].getName());
            if (a[i].getName().contains(".xlsx")) {
                System.out.println(a[i].getName());
                a[i].delete();
            }
        }
        printFileNames(a, i + 1);
    }
}
