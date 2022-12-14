package fileView;

import data.InfoList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

public class ClassTableTSV {
    Scanner scanner;

    public ClassTableTSV(File file) throws IOException, InvalidFormatException {
        scanner = new Scanner(file);
    }

    public void getClose() throws IOException {
        scanner.close();
    }

    public void getClassTable(InfoList infoList){
        scanner.nextLine();
        for(int k = 0;scanner.hasNextLine();k++){
            String line;
            line = scanner.nextLine();
            String[] elements = line.split("\t");
            if(elements[0].length() > 0){
                infoList.classTable.add(new ArrayList<>());
            }
            for (int i = 0; i < elements.length; i++){
                if(elements[i].length()>0) {
                    infoList.classTable.get(k).add(elements[i]);
                }
            }
        }
    }

}