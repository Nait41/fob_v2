package fileView;

import data.InfoList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Locale;
import java.util.Scanner;

public class PhylumTableTSV {
    Scanner scanner;

    public PhylumTableTSV(File file) throws IOException, InvalidFormatException {
        scanner = new Scanner(file);
    }

    public void getClose() throws IOException {
        scanner.close();
    }

    public void getPhylumTable(InfoList infoList){
        scanner.nextLine();
        for(int k = 0;scanner.hasNextLine();k++){
            String line;
            line = scanner.nextLine();
            String[] elements = line.split("\t");
            if(elements[0].length() > 0){
                infoList.phylumTable.add(new ArrayList<>());
            }
            for (int i = 0; i < elements.length; i++){
                if(elements[i].length()>0) {
                    infoList.phylumTable.get(k).add(elements[i]);
                }
            }
        }
    }

}