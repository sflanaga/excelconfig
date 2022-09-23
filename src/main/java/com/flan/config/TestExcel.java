package com.flan.config;

import com.flan.config.excel.Workbook;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.stream.Collectors;

public class TestExcel {

  public static void main(String[] args) {
    var times = new ArrayList<Long>();
    for (int i = 0; i < 10; i++) {
      try {
        long start = System.currentTimeMillis();
        var wb = new Workbook(Paths.get("src/main/resources/test.xlsx"));

        System.out.println(wb.sheetNames.keySet().stream().collect(Collectors.joining()));

        System.out.println(wb.tableData.keySet().stream().collect(Collectors.joining()));

        var td = wb.tableData;
        var tdbn = wb.tableDataByName;
        times.add(System.currentTimeMillis()-start);
        System.out.println((System.currentTimeMillis() - start) + " ms");

        for (var te : wb.tableDataByName.entrySet()) {
          System.out.println("table name: " + te.getKey());
          var tab = te.getValue();
          for (int row = 0; row < tab.height; row++) {
            for (int col = 0; col < tab.width; col++) {
              System.out.print(tab.data[row][col] + '\t');
            }
            System.out.println();

          }
        }
                System.out.println((System.currentTimeMillis() - start) + " ms");

      } catch (Exception e) {
        e.printStackTrace();
      }
      System.out.println("DONE");
      for(var t: times){
        System.out.println(t);
      }
    }
  }

}
