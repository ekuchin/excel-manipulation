package ru.ekuchin;

import java.io.*;

public class App
{
    public static void main( String[] args ) throws IOException {

        String filename = args[0];
        int idColumn = Integer.parseInt(args[1])-1;
        int dataColumn = Integer.parseInt(args[2])-1;

        //String filename = "/home/kea/projects/excel-manipulation/test.xlsx";
        //int idColumn = 0; int dataColumn = 4;

        System.out.println( "Идет анализ "+filename+"..." );

        Excel excel = new Excel(filename);

        String header_id=excel.getCellValueString(excel.getCell(0,idColumn));
        String header_data=excel.getCellValueString(excel.getCell(0,dataColumn));
        String curr_id = header_id;
        String curr_data = header_data;

        String prev_id;String prev_data;
        String log;
        FileWriter fileWriter = new FileWriter(filename+".report");

        for (int i=1; i<excel.getRowCount();i++){
            prev_id=curr_id;
            prev_data=curr_data;
            curr_id=excel.getCellValueString(excel.getCell(i,idColumn));
            curr_data=excel.getCellValueString(excel.getCell(i,dataColumn));
            if (curr_id.equals(prev_id) && !curr_data.equals(prev_data)){
                //<Заголовок> <id> соответствуют значения <curr_data> <prev_data>
                log = String.format("%s %s соответствуют значения %s %s и %s \n",
                        header_id, curr_id, header_data, curr_data, prev_data);
                fileWriter.write(log);
                System.out.print(log);
            }
        }
        fileWriter.close();

    }
}
