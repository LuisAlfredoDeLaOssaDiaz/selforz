import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Scanner;
/**
 *
 * @author decodigo
 */
public class Main {
    public static void main(String args[]) {
        int cantidadNotas = 6, cantidadAprobados = 0, cantidadNoAprobados = 0;
        double standardDeviation = 0;
        try {
            Scanner read = new Scanner(System.in);
            System.out.print("Ingrese ruta del archivo: ");
            String rutaArchivoExcel = read.nextLine();
            int i = 0, z = 0, cantidadEstudiantes = 0;
            String[] notas = new String[162];
            // rutaArchivoExcel = "/home/osso/Documentos/Comutacion e interfaces/semana 3/DATOS.xlsx";
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
                    contenidoCelda = contenidoCelda.replace(",", "."); // PASAR COMA A PUNTO

                    notas[i] = contenidoCelda;
                    // System.out.println(notas[i]);
                    i=i+1;
                }
            }
            cantidadEstudiantes = i/8;
            String[] guardarNotas = new String[i];
            for ( int j = 8 ; j < i ; j++ ) {
                guardarNotas[z] = notas[j];
                // System.out.println(guardarNotas[z]);
                int a = j + 9;
                if ( j == a ) {
                    j=9;
                }
                z = z + 1;
            }

            int au = 7;
            int w = 0;
            int a = 0;
            String[] onlyNotas = new String[guardarNotas.length];
            for (int g = 2 ; g < guardarNotas.length ; g++) {
                onlyNotas[w] = (guardarNotas[g]);
                // System.out.println((a++) + "  "+onlyNotas[w]);
                if (g == au) {
                    g = g + 2;
                    au = au + 8;
                }
                w++;
            }

            int f = 0;
            String filtroNotas[] = new String[guardarNotas.length];
            for (int l = 0; l < guardarNotas.length; l++) {
                if (onlyNotas[l] == null) {
                } else {
                    if (isNumeric(onlyNotas[l])) {
                    } else {
                        onlyNotas[l] = "0";
                    }
                    filtroNotas[f] = (onlyNotas[l]);
                    // System.out.println((a++) + "  "+filtroNotas[f]);
                    f++;
                }
            }
            String notasOk[] = new String[f];
            float notasOkNum[] = new float[f];
            for (int u = 0; u < f; u++) {
                if (filtroNotas[u] == null) {
                } else {
                    notasOk[u] = filtroNotas[u];
                }
                notasOkNum[u] = Float.valueOf(notasOk[u]);
                if (notasOkNum[u] < 0 || notasOkNum[u] > 5) {
                    notasOkNum[u] = 0;
                }
                // System.out.println(notasOkNum[u]);
                // System.out.println((a++) + "  "+notasOkNum[u]*2);
            }
            int cont = 0;
            String codigo[] = new String[f];
            String nombre[] = new String[f];
            for (int j = 0; j < f; j++) {
                codigo[j] = guardarNotas[cont];
                nombre[j] = guardarNotas[cont+1];
                //System.out.println(guardarNotas[cont+1]);
                cont = cont + 8;
                if (guardarNotas[cont] == null) {
                    j = guardarNotas.length + 1;
                }
            }
            int k = 0;
            double def[] = new double[cantidadEstudiantes], promedioGeneral = 0, sum = 0 ;
            for (int j = 0; j < (cantidadEstudiantes-1); j++) {
                def[j] = ((notasOkNum[k] * 0.5) + (((notasOkNum[k+1] + notasOkNum[k+2] + notasOkNum[k+3]) / 3) * 0.3) + (((notasOkNum[k+4] + notasOkNum[k+5]) / 2) * 0.2 ));// PUNTO
                //System.out.println(def[j]);
                // ((def[j] >= 3) ? cantidadAprobados++ : cantidadNoAprobados++)
                if (def[j] >= 3) {
                    cantidadAprobados++;// PUNTO
                } else {
                    cantidadNoAprobados++;// PUNTO
                }

                sum += def[j] ;

                //for (k = 0; k < f; k = k + 6 ) {
                System.out.println("Codigo : " + codigo[j] + " Nombre : " + nombre[j] + " Notas : " + notasOkNum[k] + "  " + notasOkNum[k+1] + "  " + notasOkNum[k+2] + "  " + notasOkNum[k+3] + "  " + notasOkNum[k+4] + "  " + notasOkNum[k+5] + " Definitiva : " + def[j] + ((def[j] >= 3) ? " Aprobado" : " Reprobado ") );
                //}
                k=k+cantidadNotas;
            }
            promedioGeneral = sum / cantidadEstudiantes; // PUNTO
            System.out.println("promedioGeneral : " + promedioGeneral);
            System.out.println("cantidadAprobados : " + cantidadAprobados);
            System.out.println("cantidadNoAprobados : " + cantidadNoAprobados);
            for (int n = 0; n < cantidadEstudiantes; n++) {
                standardDeviation = (standardDeviation + Math.pow((def[n] - promedioGeneral), 2)); // PUNTO
            }
            System.out.println("standardDeviation : " + standardDeviation);
            // System.out.println(standardDeviation);
            // POSICION - CODIGO - NOMBRE - DEFINITIVA - APRONOAPRO -
            String deff[] = new String[def.length];
            for (int u = 0; u < def.length; u++) {
                deff[u] = String.valueOf(def[u]);
            }
            String promedioGeneral_s;
            String standardDeviation_s;
            String cantidadAprobados_s;
            String cantidadNoAprobados_s;
            promedioGeneral_s = String.valueOf(promedioGeneral);
            standardDeviation_s = String.valueOf(standardDeviation);
            cantidadAprobados_s = String.valueOf(cantidadAprobados);
            cantidadNoAprobados_s = String.valueOf(cantidadNoAprobados);

            String exportar[][] = new String[codigo.length][codigo.length];
            for(int h = 0; i<codigo.length;h++){
                exportar[h][0] = codigo[h];
                exportar[h][1] = nombre[h];
                exportar[h][2] = String.valueOf(def[i]);
                exportar[h][3] = (def[h] >= 3.0) ? " Aprobado" : " Reprobado ";
            }
            exportar[0][4] = promedioGeneral_s;
            exportar[0][5] = standardDeviation_s;
            exportar[0][6] = cantidadAprobados_s;
            exportar[0][7] = cantidadNoAprobados_s;
            exportar(exportar);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void exportar(String[][] informacion){
        XSSFWorkbook libro= new XSSFWorkbook();
        String nombreArchivo="calificacion.xlsx";
        String directorioActual = System.getProperty("user.dir");
        String rutaArchivo= directorioActual+"\\"+nombreArchivo;
        String hoja="Resultados";
        XSSFSheet hoja1 = libro.createSheet(hoja);
        String [] cabecera= new String[]{"Codigos","Nombres","Definitivas","Aprobado/No Aprobado","Promedio general","Desviacion","Cantidad de aprobados","Cantidad de no aprobados"};
        for (int i = 0; i <= informacion.length; i++) {
            XSSFRow row=hoja1.createRow(i);//se crea las filas
            for (int j = 0; j < cabecera.length; j++) {
                XSSFCell cell= row.createCell(j);//se crea las celdas para la cabecera, junto con la posición
                //se crea las celdas para el contenido, junto con la posición
                if (i==0) {//para la cabecera
                    cell.setCellValue(cabecera[j]);//se añade el contenido
                }else{//para el contenido
                    cell.setCellValue(informacion[i-1][j]); //se añade el contenido
                }
            }
        }

        File file;
        file = new File(rutaArchivo);
        try (FileOutputStream fileOuS = new FileOutputStream(file)){
            if (file.exists()) {// si el archivo existe se elimina
                file.delete();
                System.out.println("Archivo eliminado");
            }
            libro.write(fileOuS);
            fileOuS.flush();
            fileOuS.close();
            System.out.println("Archivo Creado");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static boolean isNumeric(String str){
        return str != null && str.matches("[0-9.]+");
    }

}
