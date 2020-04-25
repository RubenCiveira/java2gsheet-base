package net.civeira.management.googlesheet;

import java.io.File;
import java.io.FileInputStream;
import java.util.Arrays;
import java.util.List;

public class Main {
  public static void main(String[] args) {
    try {
      FileInputStream fin = new FileInputStream(new File(System.getProperty("user.home") + "/Sistema/google-api/sample-api.json"));
      GoogleSheetApi sheet = new GoogleSheetApi("Civi java 2 gsheet", new GoogleAccessApi(fin), "1JhSyPRa47Fx-32lH8UQFyImBgCXCRWmOBpopHB5gCYY");
      List<List<Object>> read = sheet.read("Hoja 1" , "A", 1, "B", 4);
      for (List<Object> list : read) {
        System.out.println("Nueva Fila:");
        for (Object object : list) {
          System.out.println("\tTengo a " + object );
        }
      }
      int last = sheet.append("Hoja 1", Arrays.asList(Arrays.asList("Creado", "desde", "codigo")));
      System.out.println("Creando en la fila " + last);
      sheet.update("Hoja 1", "A", last, Arrays.asList(Arrays.asList("Updateado", "con", "java")));
      
      sheet.deleteRows("Hoja 1", 3, 4);
      sheet.deleteColumns("Hoja 1", 6, 6);
    } catch(Exception ex) {
      ex.printStackTrace();
    }
  }
}
