package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class ConvertXMLtoXSLX {
    public static void main(String[] args) throws IOException {
        Scanner scanner = new Scanner(System.in);
        System.out.print("Введите путь к папке с набором данных в формате .xml: ");
        String xmlDirectory = scanner.nextLine();
        System.out.print("Введите путь к папке, где будут сохранены XSLX файлы: ");
        String XSLXFiles = scanner.nextLine();
        System.out.print("Введите путь к шаблону XSLX: ");
        String sample = scanner.nextLine();
        File sampleXSLX = new File(sample);

        List<File> listXMLFiles = readFilesFromDir(new File(xmlDirectory), ".xml");

        for (File xml : listXMLFiles) {
            String nameFile = xml.getName().substring(0, xml.getName().length() - 4);
            List<CatalogSeqNumber> catalogSeqNumbers = getParseXML(xml);
            createExcelFile(catalogSeqNumbers, XSLXFiles, nameFile, sampleXSLX);
        }


    }


    private static void createExcelFile(List<CatalogSeqNumber> catalogSeqNumbers, String outputFilePath, String nameFile, File sampleXSLX) throws IOException {
        try {
            FileInputStream fis = new FileInputStream(sampleXSLX);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);

            int rowNum = 2;

            for (CatalogSeqNumber catalogSeqNumber : catalogSeqNumbers) {
                for (ItemSeqNumber itemSeqNumber : catalogSeqNumber.getItemSeqNumbers()) {
                    for (PartRef partRef : itemSeqNumber.getPartRefs()) {
                        Row row = sheet.createRow(rowNum++);

                        row.createCell(0).setCellValue(partRef.partNumberValue);
                        row.createCell(1).setCellValue(catalogSeqNumber.getItem());
                        row.createCell(16).setCellValue(itemSeqNumber.getItemSeqNumberValue());
                        if (!itemSeqNumber.getDescrForParts().isEmpty()) {
                            row.createCell(3).setCellValue(itemSeqNumber.getDescrForParts().get(0).getDescrForPart());
                        }
                        if (!itemSeqNumber.getGenericPartDataValues().isEmpty()) {
                            row.createCell(4).setCellValue(itemSeqNumber.getGenericPartDataValues().get(0).getGenericPartDataValue());
                        }
                        if (!itemSeqNumber.getDescrForLocations().isEmpty()) {
                            row.createCell(5).setCellValue(itemSeqNumber.getDescrForLocations().get(0).getDescrForLocationValue());
                        }

                        row.createCell(6).setCellValue(1);

                        if (!itemSeqNumber.getNotIllustrates().isEmpty()) {
                            if(itemSeqNumber.getNotIllustrates().get(0).getNotIllustratedValue().equals("-")){
                                row.createCell(22).setCellValue(1);
                            }

                        }

                        row.createCell(41).setCellValue(itemSeqNumber.getApplicRefIds());

                    }
                }
            }

            // Сохраняем изменения
            try (FileOutputStream fos = new FileOutputStream(outputFilePath + "\\" + nameFile + ".xlsx")) {
                workbook.write(fos);
            }
//"C:\Users\SurfaceBook\Desktop\XMLtoXSLX\sample.xlsx"
            // Закрываем рабочую книгу
            workbook.close();
            fis.close();

            System.out.println("Данные успешно изменены в файле: " + outputFilePath + nameFile);
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private static List<CatalogSeqNumber> getParseXML(File xml) throws IOException {

        List<CatalogSeqNumber> catalogSeqNumbers = new ArrayList<>();
        try {
            //FileInputStream fileInputStream = new FileInputStream(sampleXSLX);
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.parse(xml);
            doc.getDocumentElement().normalize();

            //Получение применяемости

            // Получаем список всех элементов <applic>
            NodeList applicList = doc.getElementsByTagName("applic");

            // Создаем словарь
            Map<String, String> applicMap = new HashMap<>();

            // Проходим по всем элементам <applic>
            for (int i = 0; i < applicList.getLength(); i++) {
                Node applicNode = applicList.item(i);
                if (applicNode.getNodeType() == Node.ELEMENT_NODE) {
                    Element applicElement = (Element) applicNode;

                    // Получаем id
                    String id = applicElement.getAttribute("id");

                    // Получаем элемент <assert>
                    NodeList assertList = applicElement.getElementsByTagName("assert");
                    if (assertList.getLength() > 0) {
                        Element assertElement = (Element) assertList.item(0);

                        // Получаем значение applicPropertyValues
                        String applicPropertyValues = assertElement.getAttribute("applicPropertyValues");

                        // Добавляем в словарь
                        applicMap.put(id, applicPropertyValues);
                    }
                }
            }

            //Извлечение списка catalogSeqNumber
            NodeList nodeList = doc.getElementsByTagName("catalogSeqNumber");

            //Перебор catalogSeqNumber
            for (int i = 0; i < nodeList.getLength(); i++) {
                Node node = nodeList.item(i);
                if (node.getNodeType() == Node.ELEMENT_NODE) {
                    Element catalogSeqNumberElement = (Element) node;
                    // Извлечение атрибутов catalogSeqNumber
                    CatalogSeqNumber catalogSeqNumber = new CatalogSeqNumber();
                    catalogSeqNumber.setSystemCode(catalogSeqNumberElement.getAttribute("systemCode"));
                    catalogSeqNumber.setSubSystemCode(catalogSeqNumberElement.getAttribute("subSubSystemCode"));
                    catalogSeqNumber.setAssyCode(catalogSeqNumberElement.getAttribute("assyCode"));
                    catalogSeqNumber.setFigureNumber(catalogSeqNumberElement.getAttribute("figureNumber"));
                    catalogSeqNumber.setItem(catalogSeqNumberElement.getAttribute("item"));
                    catalogSeqNumber.setIndenture(catalogSeqNumberElement.getAttribute("indenture"));
                    catalogSeqNumber.setFigureNumberVariant(catalogSeqNumberElement.getAttribute("figureNumberVariant"));

                    // Получение списка элементов itemSeqNumber
                    NodeList itemSeqNumberList = catalogSeqNumberElement.getElementsByTagName("itemSeqNumber");
                    List<ItemSeqNumber> itemSeqNumbers = new ArrayList<>();

                    // Перебор элементов itemSeqNumber
                    for (int j = 0; j < itemSeqNumberList.getLength(); j++) {
                        Node itemSeqNumberNode = itemSeqNumberList.item(j);
                        if (itemSeqNumberNode.getNodeType() == Node.ELEMENT_NODE) {

                            Element itemSeqNumberElement = (Element) itemSeqNumberNode;

                            ItemSeqNumber itemSeqNumber = new ItemSeqNumber();
                            itemSeqNumber.setItemSeqNumberValue(itemSeqNumberElement.getAttribute("itemSeqNumberValue"));


                            itemSeqNumber.setApplicRefIds(applicMap.get(itemSeqNumberElement.getAttribute("applicRefIds")));

                            List<PartRef> partRefs = new ArrayList<>();
                            List<DescrForPart> descrForParts = new ArrayList<>();
                            List<GenericPartDataValue> genericPartDataValues = new ArrayList<>();
                            List<DescrForLocation> descrForLocationsList = new ArrayList<>();
                            List<NotIllustrated> notIllustratedList = new ArrayList<>();

                            // Извлечение элементов partRef
                            NodeList partRefList = itemSeqNumberElement.getElementsByTagName("partRef");
                            for (int k = 0; k < partRefList.getLength(); k++) {
                                Node partRefNode = partRefList.item(k);
                                if (partRefNode.getNodeType() == Node.ELEMENT_NODE) {
                                    Element partRefElement = (Element) partRefNode;

                                    PartRef partRef = new PartRef();
                                    partRef.setManufacturerCodeValue(partRefElement.getAttribute("manufacturerCodeValue"));
                                    partRef.setPartNumberValue(partRefElement.getAttribute("partNumberValue"));
                                    partRefs.add(partRef);
                                }
                            }

                            // Извлечение элементов descrForPart
                            NodeList descrForPartList = itemSeqNumberElement.getElementsByTagName("descrForPart");
                            for (int k = 0; k < descrForPartList.getLength(); k++) {
                                Node descrForPartNode = descrForPartList.item(k);
                                if (descrForPartNode.getNodeType() == Node.ELEMENT_NODE) {
                                    Element descrForPartElement = (Element) descrForPartNode;

                                    DescrForPart descrForPart = new DescrForPart();
                                    descrForPart.setDescrForPart(descrForPartElement.getTextContent());
                                    descrForParts.add(descrForPart);
                                }
                            }

                            // Извлечение элементов genericPartDataValue
                            NodeList genericPartDataValueList = itemSeqNumberElement.getElementsByTagName("genericPartDataValue");
                            for (int k = 0; k < genericPartDataValueList.getLength(); k++) {
                                Node genericPartDataValueNode = genericPartDataValueList.item(k);
                                if (genericPartDataValueNode.getNodeType() == Node.ELEMENT_NODE) {
                                    Element genericPartDataValueElement = (Element) genericPartDataValueNode;

                                    GenericPartDataValue genericPartDataValue = new GenericPartDataValue();
                                    genericPartDataValue.setGenericPartDataValue(genericPartDataValueElement.getTextContent());
                                    genericPartDataValues.add(genericPartDataValue);
                                }
                            }

                            // Извлечение элементов notIllustrated
                            NodeList notIllustrated = itemSeqNumberElement.getElementsByTagName("notIllustrated");
                            for (int k = 0; k < notIllustrated.getLength(); k++) {
                                Node notIllustratedNode = notIllustrated.item(k);
                                if (notIllustratedNode.getNodeType() == Node.ELEMENT_NODE) {
                                    Element notIllustratedNodeElement = (Element) notIllustratedNode;

                                    NotIllustrated notIllustratedclass = new NotIllustrated();
                                    notIllustratedclass.setNotIllustratedValue(notIllustratedNodeElement.getTextContent());
                                    notIllustratedList.add(notIllustratedclass);
                                }
                            }
                            // Извлечение элементов descrForLocation

                            NodeList descrForLocationList = itemSeqNumberElement.getElementsByTagName("descrForLocation");
                            for (int k = 0; k < descrForLocationList.getLength(); k++) {
                                Node descrForLocationNode = descrForLocationList.item(k);
                                if (descrForLocationNode.getNodeType() == Node.ELEMENT_NODE) {
                                    Element descrForLocationNodeElement = (Element) descrForLocationNode;

                                    DescrForLocation DescrForLocationclass = new DescrForLocation();
                                    DescrForLocationclass.setDescrForLocationValue(descrForLocationNodeElement.getTextContent());
                                    descrForLocationsList.add(DescrForLocationclass);
                                }
                            }

                            itemSeqNumber.setPartRefs(partRefs);
                            itemSeqNumber.setDescrForParts(descrForParts);
                            itemSeqNumber.setGenericPartDataValues(genericPartDataValues);
                            itemSeqNumber.setNotIllustrates(notIllustratedList);
                            itemSeqNumber.setDescrForLocations(descrForLocationsList);
                            itemSeqNumbers.add(itemSeqNumber);
                        }

                    }

                    catalogSeqNumber.setItemSeqNumbers(itemSeqNumbers);
                    catalogSeqNumbers.add(catalogSeqNumber);

                }

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return catalogSeqNumbers;
    }


    private static List<File> readFilesFromDir(File xmlDirectory, String ext) {
        List<File> files = new ArrayList<>();
        for (File f : Objects.requireNonNull(xmlDirectory.listFiles())) {
            if (f.isFile() && f.getName().toLowerCase().endsWith(ext)) files.add(f);
            else if (f.isDirectory()) {
                List<File> dirFiles = readFilesFromDir(f, ext);
                files.addAll(dirFiles);
            }
        }
        return files;

    }

    public static class CatalogSeqNumber {
        private String systemCode;
        private String subSystemCode;
        private String subSubSystemCode;
        private String assyCode;
        private String figureNumber;
        private String item;
        private String indenture;
        private String figureNumberVariant;
        private List<ItemSeqNumber> itemSeqNumbers;

        // Геттеры и сеттеры

        public String getSystemCode() {
            return systemCode;
        }

        public void setSystemCode(String systemCode) {
            this.systemCode = systemCode;
        }

        public String getSubSystemCode() {
            return subSystemCode;
        }

        public void setSubSystemCode(String subSystemCode) {
            this.subSystemCode = subSystemCode;
        }

        public String getSubSubSystemCode() {
            return subSubSystemCode;
        }

        public void setSubSubSystemCode(String subSubSystemCode) {
            this.subSubSystemCode = subSubSystemCode;
        }

        public String getAssyCode() {
            return assyCode;
        }

        public void setAssyCode(String assyCode) {
            this.assyCode = assyCode;
        }

        public String getFigureNumber() {
            return figureNumber;
        }

        public void setFigureNumber(String figureNumber) {
            this.figureNumber = figureNumber;
        }

        public String getItem() {
            return item;
        }

        public void setItem(String item) {
            this.item = item;
        }

        public String getIndenture() {
            return indenture;
        }

        public void setIndenture(String indenture) {
            this.indenture = indenture;
        }

        public String getFigureNumberVariant() {
            return figureNumberVariant;
        }

        public void setFigureNumberVariant(String figureNumberVariant) {
            this.figureNumberVariant = figureNumberVariant;
        }

        public List<ItemSeqNumber> getItemSeqNumbers() {
            return itemSeqNumbers;
        }

        public void setItemSeqNumbers(List<ItemSeqNumber> itemSeqNumbers) {
            this.itemSeqNumbers = itemSeqNumbers;
        }
    }

    public static class ItemSeqNumber {
        private String itemSeqNumberValue;
        private String applicRefIds;
        private List<PartRef> partRefs;
        private List<DescrForPart> descrForParts;
        private List<GenericPartDataValue> genericPartDataValues;
        private List<NotIllustrated> notIllustrates;
        private  List<DescrForLocation> descrForLocations;

        public List<NotIllustrated> getNotIllustrates() {
            return notIllustrates;
        }

        public void setNotIllustrates(List<NotIllustrated> notIllustrates) {
            this.notIllustrates = notIllustrates;
        }

        public List<DescrForLocation> getDescrForLocations() {
            return descrForLocations;
        }

        public void setDescrForLocations(List<DescrForLocation> descrForLocations) {
            this.descrForLocations = descrForLocations;
        }
// Геттеры и сеттеры

        public String getItemSeqNumberValue() {
            return itemSeqNumberValue;
        }

        public void setItemSeqNumberValue(String itemSeqNumberValue) {
            this.itemSeqNumberValue = itemSeqNumberValue;
        }

        public String getApplicRefIds() {
            return applicRefIds;
        }

        public void setApplicRefIds(String applicRefIds) {
            this.applicRefIds = applicRefIds;
        }

        public List<PartRef> getPartRefs() {
            return partRefs;
        }

        public void setPartRefs(List<PartRef> partRefs) {
            this.partRefs = partRefs;
        }

        public List<DescrForPart> getDescrForParts() {
            return descrForParts;
        }

        public void setDescrForParts(List<DescrForPart> descrForParts) {
            this.descrForParts = descrForParts;
        }

        public List<GenericPartDataValue> getGenericPartDataValues() {
            return genericPartDataValues;
        }

        public void setGenericPartDataValues(List<GenericPartDataValue> genericPartDataValues) {
            this.genericPartDataValues = genericPartDataValues;
        }
    }

    public static class PartRef {
        private String manufacturerCodeValue;
        private String partNumberValue;

        // Геттеры и сеттеры


        public String getManufacturerCodeValue() {
            return manufacturerCodeValue;
        }

        public void setManufacturerCodeValue(String manufacturerCodeValue) {
            this.manufacturerCodeValue = manufacturerCodeValue;
        }

        public String getPartNumberValue() {
            return partNumberValue;
        }

        public void setPartNumberValue(String partNumberValue) {
            this.partNumberValue = partNumberValue;
        }
    }

    public static class DescrForPart {
        private String descrForPart;

        // Геттеры и сеттеры

        public String getDescrForPart() {
            return descrForPart;
        }

        public void setDescrForPart(String descrForPart) {
            this.descrForPart = descrForPart;
        }
    }

    public static class GenericPartDataValue {
        private String genericPartDataValue;

        // Геттеры и сеттеры

        public String getGenericPartDataValue() {
            return genericPartDataValue;
        }

        public void setGenericPartDataValue(String genericPartDataValue) {
            this.genericPartDataValue = genericPartDataValue;
        }
    }

    public static class NotIllustrated {
        private String notIllustratedValue;

        public String getNotIllustratedValue() {
            return notIllustratedValue;
        }

        public void setNotIllustratedValue(String notIllustratedValue) {
            this.notIllustratedValue = notIllustratedValue;
        }

    }

    public static class DescrForLocation {
        private String descrForLocationValue;

        public String getDescrForLocationValue() {
            return descrForLocationValue;
        }

        public void setDescrForLocationValue(String descrForLocationValue) {
            this.descrForLocationValue = descrForLocationValue;
        }
    }
}






