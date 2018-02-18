package com.company;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.*;

import javax.xml.crypto.Data;
import java.io.*;
import java.util.Scanner;

public class Main {


    public static void main(String[] args) throws IOException {


        int a=0,b,c=0,d=0,f;
        double num;
        String name,address1,address2;
        Scanner scin = new Scanner(System.in);

        System.out.println("Введите путь Excel файла");
        address1=scin.nextLine();
        Workbook wb = new HSSFWorkbook(new FileInputStream(address1+".xls"));

        System.out.println("Введите путь сохранения");
        address2=scin.nextLine();
        FileWriter fw= new FileWriter(address2+".mkb");
        System.out.println("Введите название базы знаний");
        name=scin.nextLine();
        fw.write(name);
        fw.write(System.getProperty("line.separator"));
        System.out.println("Введите автора");
        name=scin.nextLine();

        fw.write("Автор:"+name);
        fw.write(System.getProperty("line.separator"));


        Sheet wbSheet =  wb.getSheetAt(0);
        Row wbRow=wbSheet.getRow(1);


        for(Row row:wbSheet){
           a++; }


           for(Cell cell:wbRow){
            name = cell.getRichStringCellValue().getString();
               if(!name.equals(""))
                   fw.write(System.getProperty("line.separator"));
               fw.write(name);
               if(d==0){fw.write(":");}
               d++;

        }
        fw.write(System.getProperty("line.separator"));
        fw.write(System.getProperty("line.separator"));
        for(int i=3;i!=(a-1);i++) {
            wbRow=wbSheet.getRow(i);
            c=1;
            b=0;

                for (Cell cell :wbRow){

                    if ((b%2==0&&b>0)){fw.write(c+",");c++;}
                    switch (cell.getCellTypeEnum()) {
                        case STRING:
                            name=cell.getStringCellValue();
                                fw.write(name+",");
                            break;
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                name=cell.getDateCellValue().toString();
                                fw.write(name+",");
                            } else {
                               num=cell.getNumericCellValue();
                                fw.write(num+",");
                            }
                            break;
                        case BOOLEAN:
                            System.out.println(cell.getBooleanCellValue());
                            break;
                        case FORMULA:
                            num=cell.getNumericCellValue();
                            fw.write(num+"");
                            if(b<(d)){
                            fw.write(",");}
                            break;
                        case BLANK:
                            System.out.println();
                            break;
                        default:
                            System.out.println();

                    }b++;

                }

            if (i<(a-2)){fw.write(System.getProperty("line.separator"));}
        }





        fw.close();

    }
}
