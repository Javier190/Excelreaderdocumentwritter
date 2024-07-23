package org.example;

import org.example.domain.Person;
import org.example.util.ExcelReader;
import org.example.util.WordCreator;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.List;

public class Main {

    public static void main(String[] args) {
        JFrame frame = new JFrame("Generador de Documentos Word");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(400, 300);
        frame.setLayout(new GridLayout(5, 1));

        JPanel panelExcel = new JPanel();
        panelExcel.add(new JLabel("Ruta del archivo Excel:"));
        JTextField excelPathField = new JTextField(20);
        panelExcel.add(excelPathField);

        JPanel panelTemplate = new JPanel();
        panelTemplate.add(new JLabel("Ruta de la plantilla de Word:"));
        JTextField templatePathField = new JTextField(20);
        panelTemplate.add(templatePathField);

        JPanel panelOutput = new JPanel();
        panelOutput.add(new JLabel("Directorio de salida:"));
        JTextField outputDirField = new JTextField(20);
        panelOutput.add(outputDirField);

        JButton executeButton = new JButton("Generar Documentos");

        JTextArea outputArea = new JTextArea(5, 30);
        outputArea.setEditable(false);
        JScrollPane scrollPane = new JScrollPane(outputArea);

        executeButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String excelFilePath = excelPathField.getText();
                String templateFilePath = templatePathField.getText();
                String outputDirectory = outputDirField.getText();

                ExcelReader excelReader = new ExcelReader();
                WordCreator wordCreator = new WordCreator();

                try {
                    List<Person> people = excelReader.readExcel(excelFilePath);
                    wordCreator.createWordDocuments(people, templateFilePath, outputDirectory);
                    outputArea.append("¡Documentos de Word creados exitosamente!\n");
                } catch (Exception ex) {
                    outputArea.append("Error: " + ex.getMessage() + "\n");
                    ex.printStackTrace();
                }
            }
        });

        frame.add(panelExcel);
        frame.add(panelTemplate);
        frame.add(panelOutput);
        frame.add(executeButton);
        frame.add(scrollPane);

        frame.setVisible(true);
    }
}
