import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *
 * @author decodigo
 */
public class Main {
    public static void main(String args[]) {
        try {
            int i = 0, z = 0;
            String[] notas = new String[162];
            String rutaArchivoExcel = "/home/osso/Documentos/Comutacion e interfaces/semana 3/DATOS.xlsx";
            FileInputStream inputStream = new FileInputStream(new File(rutaArchivoExcel));
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet firstSheet = workbook.getSheetAt(0);
            Iterator iterator = firstSheet.iterator();

            DataFormatter formatter = new DataFormatter();

            while (iterator.hasNext()) {

                Row nextRow = (Row) iterator.next();

                Iterator cellIterator = nextRow.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = (Cell) cellIterator.next();
                    String contenidoCelda = formatter.formatCellValue(cell);

                    contenidoCelda = ((contenidoCelda == "" ) ? "0" : contenidoCelda );
                    contenidoCelda = contenidoCelda.replace(",", ".");


                    notas[i] = contenidoCelda;
                    i=i+1;
                }
            }

            String[] guardarNotas = new String[i];

            for ( int j = 9 ; j < i ; j++ ) {
                guardarNotas[z] = notas[j];
                int a = j + 9;
                if ( j == a ) {
                    j=9;
                }
                z = z + 1;
            }

            int au = 7;
            for (int g = 2 ; g < guardarNotas.length ; g++) {

                if (g == au) {
                   g = g + 3;
                   au = au + 7;
               }
                System.out.println(guardarNotas[g]);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static boolean isNumeric(String str){
        return str != null && str.matches("[0-9.]+");
    }
}
