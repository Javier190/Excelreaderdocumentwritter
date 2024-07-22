package org.example;

import org.example.domain.Person;
import org.example.util.ExcelReader;
import org.example.util.WordCreator;


import java.util.List;


public class Main {
    public static void main(String[] args) {
        String excelFilePath = "C:\\Users\\javar\\Desktop\\personas.xlsx";  // Ruta del archivo Excel
        String templateFilePath = "C:\\Users\\javar\\Desktop\\template.docx";  // Ruta de la plantilla de Word
        String outputDirectory = "C:\\Users\\javar\\Desktop";  // Directorio de salida (ruta correcta)

        ExcelReader excelReader = new ExcelReader();
        WordCreator wordCreator = new WordCreator();

        try {
            List<Person> people = excelReader.readExcel(excelFilePath);
            wordCreator.createWordDocuments(people, templateFilePath, outputDirectory);
            System.out.println("Word documents created successfully!");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

