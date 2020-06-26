package exports;

import com.fasterxml.jackson.annotation.JsonFormat;
import com.fasterxml.jackson.annotation.JsonIgnore;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.nio.charset.StandardCharsets;
import java.sql.Time;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class Exports<T> {
    public String CSVSeparator = ";";
    public String language = "pt";
    public String country = "BR";


    private String generateRandomFile(String extensao){
        new File(System.getProperty("user.dir") + File.separatorChar + "TMP"+File.separatorChar).mkdir();

        String path = System.getProperty("user.dir") + File.separatorChar + "TMP"+File.separatorChar;
        Random random = new Random();

        String fileLocation = path + "export_"+Integer.toString( Math.abs(random.nextInt()));
        if (extensao.startsWith("."))
            fileLocation = fileLocation + extensao;
        else
            fileLocation= fileLocation + "."+extensao;
        return fileLocation;
    }

    private String valueToStr(T classe, Field field ) throws IllegalAccessException, IOException, ParseException {

        field.setAccessible(true);
        try {
            Object obj = field.get(classe);

            if (obj == null)
                return "null";
            else {
                if (field.getAnnotation(JsonFormat.class) != null) {
                    JsonFormat custom = field.getAnnotation(JsonFormat.class);
                    if ((field.getType().equals(Date.class)) || (field.getType().equals(Time.class))) {

                        Locale local = null;
                        if ((custom.locale() != null) && (custom.locale().isEmpty()))
                            local = new Locale(this.language, this.country);
                        else if ((custom.locale() != null) && (!custom.locale().isEmpty()))
                            local = new Locale(custom.locale());

                        SimpleDateFormat fmt = new SimpleDateFormat(custom.pattern(), local);
                        return fmt.format(obj);
                    } else if ((field.getType().equals(Number.class)) || (field.getType().equals(Float.class))) {
                        DecimalFormat fmtNumber = new DecimalFormat(custom.pattern());
                        return fmtNumber.format(obj);
                    } else
                        return obj.toString();
                } else if (field.getAnnotation(ExportsFromTo.class) != null) {
                    ExportsFromTo custom = field.getAnnotation(ExportsFromTo.class);
                    List<String> listFrom = Arrays.asList(custom.from());
                    List<String> listTo = Arrays.asList(custom.to());
                    String retorno = "";
                    if (listFrom.contains(obj.toString())) {
                        int item = 0;
                        for (String i : listFrom){
                            if (i.equals(obj.toString())) {
                                retorno = listTo.get(item);
                                break;
                            }
                            item++;
                        }
                        if (retorno.isEmpty()) {
                            retorno = obj.toString();
                        }
                        return retorno;
                    } else
                        return obj.toString();
                } else {
                    return obj.toString();
                }
            }
        } finally {
            field.setAccessible(false);
        }
    }

    private String nameToStr(T classe, Field field ) {

        String retorno;
        byte[] b;
        if (field.getAnnotation(ExportsName.class) != null) {
            ExportsName custom = field.getAnnotation(ExportsName.class);
            b  = custom.name().getBytes();
            retorno = new String(b,StandardCharsets.UTF_8);
        } else {
            retorno = field.getName();
        }
        return retorno;
    }

    public String exportListToCSV(List<T> lista) throws IllegalAccessException, IOException, ParseException {

        OutputStream output = null;
        String arquivo = generateRandomFile(".csv");

        try {
            output = new FileOutputStream(arquivo);

            //Cabeçalho
            String cabecalho = "";
            for (T item : lista) {
                Field[] fields = item.getClass().getDeclaredFields();
                for (Field f : fields) {
                    if( (f.getAnnotation(JsonIgnore.class) == null) && (f.getAnnotation(ExportsIgnore.class) == null) )
                        cabecalho = cabecalho + nameToStr(item,f) + this.CSVSeparator;
                }
                cabecalho = cabecalho.substring(0,cabecalho.length()-1)+"\n";
                output.write(cabecalho.getBytes());
                break;
            }
            //Valores
            StringBuffer valores = new StringBuffer();
            for (T item : lista) {
                valores.setLength(0);
                Field[] fields = item.getClass().getDeclaredFields();
                for (Field f : fields) {
                    if( (f.getAnnotation(JsonIgnore.class) == null) && (f.getAnnotation(ExportsIgnore.class) == null) )
                        valores.append(valueToStr(item, f) + this.CSVSeparator);
                }
                valores.delete(valores.length()-1,valores.length());
                valores.append("\n");
                output.write(valores.toString().getBytes());
            }
            valores.setLength(0);
        }
        finally {
            output.flush();
            output.close();
        }
        return arquivo;
    }

    public String exportListToExcel(List<T> lista, String nomePlanilha) throws IllegalAccessException, IOException, ParseException {
        Workbook workbook = new XSSFWorkbook();
        try {
            Sheet sheet = workbook.createSheet(nomePlanilha);
            //sheet.setColumnWidth(0, 6000);
            //sheet.setColumnWidth(1, 4000);

            Row header = sheet.createRow(0);

            CellStyle headerStyle = workbook.createCellStyle();
            //headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
            //headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            XSSFFont fontCabelho = ((XSSFWorkbook) workbook).createFont();
            fontCabelho.setFontName("Arial");
            fontCabelho.setFontHeightInPoints((short) 12);
            fontCabelho.setBold(true);
            headerStyle.setFont(fontCabelho);

            int linha = 0; int coluna = 0;
            /*  Cabeçalho */
            Cell headerCell;
            for (T item : lista) {
                Field[] fields = item.getClass().getDeclaredFields();
                for (Field f : fields) {
                    if( (f.getAnnotation(JsonIgnore.class) != null) || (f.getAnnotation(ExportsIgnore.class) != null) )
                        continue;

                    headerCell = header.createCell(coluna);
                    headerCell.setCellValue(nameToStr(item,f));
                    headerCell.setCellStyle(headerStyle);
                    coluna++;
                }
                break;
            }

            //Valores
            Row row = null;
            Cell cell = null;
            linha = 0;
            coluna = 0;
            for (T item : lista) {
                Field[] fields = item.getClass().getDeclaredFields();

                linha++;
                row = sheet.createRow(linha);

                coluna = 0;
                for (Field f : fields) {
                    if( (f.getAnnotation(JsonIgnore.class) != null) || (f.getAnnotation(ExportsIgnore.class) != null) )
                        continue;
                    cell = row.createCell(coluna);
                    cell.setCellValue(valueToStr(item, f));

                    coluna++;
                }
            }

            String arquivo = generateRandomFile(".xlsx");

            FileOutputStream outputStream = new FileOutputStream(arquivo);
            workbook.write(outputStream);

            return arquivo;

        } finally {
            workbook.close();
        }

    }
}
