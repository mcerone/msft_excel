import com.murilo.excel.ExcelHandler;
import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;
/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author munascim
 */
public class CsvGenerator {
    
        
    public static void generateCSV(String inputFile,String outputFile) throws IOException{
        ExcelHandler eh = new ExcelHandler(inputFile);
        BufferedWriter bw = new BufferedWriter(new FileWriter(outputFile));
        int limit = eh.getCountRowsFromSheet(0);
        String linha = "";
        int realLineCounter = 0;
        for (int i = 0; i < limit; i++) {
            linha = eh.getLine(0, i,',');
            if (linha.trim().equals("")) {
                continue;
            }
            bw.append(linha);
            realLineCounter++;
            System.out.println(realLineCounter+": "+linha);
        }
        bw.flush();
        
    }
    
    
}
