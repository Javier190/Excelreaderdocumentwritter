package org.example.util;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Text;
import org.example.domain.Person;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import java.io.File;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class WordCreator {
    public void createWordDocuments(List<Person> people, String templatePath, String outputDir) throws Exception {
        for (Person person : people) {
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(templatePath));
            MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
            Map<String, String> mappings = new HashMap<>();
            mappings.put("{{firstName}}", person.getFirstName());
            mappings.put("{{lastName}}", person.getLastName());
            mappings.put("{{rut}}", person.getRut());

            replacePlaceholders(documentPart, mappings);
            String outputFilePath = outputDir + "\\" + person.getFirstName() + "_" + person.getLastName() + ".docx";
            wordMLPackage.save(new File(outputFilePath));
        }
    }

    private void replacePlaceholders(MainDocumentPart documentPart, Map<String, String> mappings) throws Docx4JException, JAXBException {
        List<Object> texts = documentPart.getJAXBNodesViaXPath("//w:t", true);
        StringBuilder combinedText = new StringBuilder();

        for (Object obj : texts) {
            Text text;
            if (obj instanceof JAXBElement) {
                JAXBElement<?> element = (JAXBElement<?>) obj;
                if (element.getValue() instanceof Text) {
                    text = (Text) element.getValue();
                } else {
                    continue;
                }
            } else if (obj instanceof Text) {
                text = (Text) obj;
            } else {
                continue;
            }
            combinedText.append(text.getValue());
        }

        String textValue = combinedText.toString();
        for (Map.Entry<String, String> entry : mappings.entrySet()) {
            textValue = textValue.replace(entry.getKey(), entry.getValue());
        }

        int startIndex = 0;
        for (Object obj : texts) {
            Text text;
            if (obj instanceof JAXBElement) {
                JAXBElement<?> element = (JAXBElement<?>) obj;
                if (element.getValue() instanceof Text) {
                    text = (Text) element.getValue();
                } else {
                    continue;
                }
            } else if (obj instanceof Text) {
                text = (Text) obj;
            } else {
                continue;
            }

            int endIndex = Math.min(startIndex + text.getValue().length(), textValue.length());
            text.setValue(textValue.substring(startIndex, endIndex));
            startIndex = endIndex;
        }
    }
}
