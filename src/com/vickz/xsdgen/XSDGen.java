package com.vickz.xsdgen;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;
import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.XMLWriter;

public class XSDGen {
    private static final String XSD_PREFIX = "xsd";

    private Document documentMain;
    private Document documentStructure;
    private String mainXSDFileName;
    private String structureXSDFileName;

    public XSDGen parse(File excelFile) {
        parseDocumentMain(excelFile);
        parseDocumentStructure(excelFile);

        setXsdFileName(excelFile.getName());
        return this;
    }

    private void parseDocumentMain(File excelFile) {
        documentMain = DocumentHelper.createDocument();
        Element root = initialDocument(documentMain);

        XSSFWorkbook workbook;
        try {
            workbook = new XSSFWorkbook(new FileInputStream(excelFile));
            XSSFSheet mainSheet = workbook.getSheet("案件所有表");
            XSSFRow row = mainSheet.getRow(0);
            XSSFCell cell = row.getCell(1);
            Element element = root.addElement(XSD_PREFIX + ":complexType")
                    .addAttribute("name", cell.getStringCellValue());
            element = element.addElement(XSD_PREFIX + ":sequence");
            root.addElement(XSD_PREFIX + ":element")
                    .addAttribute("name", cell.getStringCellValue())
                    .addAttribute("type", cell.getStringCellValue());

            int rows = mainSheet.getPhysicalNumberOfRows();
            for (int i = 1; i < rows; i++) {
                row = mainSheet.getRow(i);
                String tableName = row.getCell(0).getStringCellValue();
                String tableChsName = row.getCell(1).getStringCellValue();
                element.addElement(XSD_PREFIX + ":element")
                        .addAttribute("name", tableName)
                        .addAttribute("type", tableChsName)
                        .addAttribute("minOccurs", "0")
                        .addAttribute("maxOccurs", "1");
            }
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

    }

    private void parseDocumentStructure(File excelFile) {
        documentStructure = DocumentHelper.createDocument();
        Element root = initialDocument(documentStructure);
        root.addElement(XSD_PREFIX + ":include").addAttribute("schemaLocation",
                "_A_类型_基础.xsd");

        XSSFWorkbook workbook;
        try {
            workbook = new XSSFWorkbook(new FileInputStream(excelFile));
            XSSFSheet structureSheet = workbook.getSheet("详细表结构");

            int rows = structureSheet.getPhysicalNumberOfRows();

            XSSFRow row = structureSheet.getRow(1);
            XSSFCell cell = row.getCell(0);
            String tableChsName = cell.getStringCellValue();
            Element element = root.addElement(XSD_PREFIX + ":complexType");

            if (row.getCell(1) != null
                    && row.getCell(1).getStringCellValue().equals("R")) {
                element.addAttribute("name", tableChsName + '1');
                root.addElement(XSD_PREFIX + ":complexType")
                        .addAttribute("name", tableChsName)
                        .addElement(XSD_PREFIX + ":sequence")
                        .addElement(XSD_PREFIX + ":element")
                        .addAttribute("name", "R")
                        .addAttribute("type", tableChsName + "1")
                        .addAttribute("minOccurs", "0")
                        .addAttribute("maxOccurs", "1");
            } else {
                element.addAttribute("name", tableChsName);
            }
            element = element.addElement(XSD_PREFIX + ":sequence");

            for (int i = 2; i < rows; i++) {
                row = structureSheet.getRow(i);
                cell = row.getCell(0);
                if (XSSFCell.CELL_TYPE_STRING == cell.getCellType()) {
                    tableChsName = cell.getStringCellValue();
                    element = root.addElement(XSD_PREFIX + ":complexType");
                    if (row.getCell(1) != null
                            && row.getCell(1).getStringCellValue().equals("R")) {
                        element.addAttribute("name", tableChsName + '1');
                        root.addElement(XSD_PREFIX + ":complexType")
                                .addAttribute("name", tableChsName)
                                .addElement(XSD_PREFIX + ":sequence")
                                .addElement(XSD_PREFIX + ":element")
                                .addAttribute("name", "R")
                                .addAttribute("type", tableChsName + "1")
                                .addAttribute("minOccurs", "0")
                                .addAttribute("maxOccurs", "1");
                    } else {
                        element.addAttribute("name", tableChsName);
                    }
                    element = element.addElement(XSD_PREFIX + ":sequence");
                } else {
                    element.addElement(XSD_PREFIX + ":element")
                            .addAttribute("name",
                                    row.getCell(2).getStringCellValue())
                            .addAttribute("type",
                                    row.getCell(3).getStringCellValue())
                            .addAttribute("minOccurs", "0")
                            .addAttribute("maxOccurs", "1");
                }

            }

        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    private Element initialDocument(Document document) {
        Element element = document.addElement(XSD_PREFIX + ":schema",
                "http://www.w3.org/2001/XMLSchema");
//        element.addNamespace(XSD_PREFIX, "http://dataexchange.court.gov.cn/2009/data");
        element.addAttribute("targetNamespace",
                "http://dataexchange.court.gov.cn/2009/data");
        element.addAttribute("elementFormDefault", "qualified");
        element.addElement(XSD_PREFIX + ":include").addAttribute(
                "schemaLocation", "_0_结构_复用.xsd");
        return element;
    }

    private void setXsdFileName(String excelFileName) {
        this.mainXSDFileName = excelFileName.substring(0,
                excelFileName.indexOf('.'))
                + ".xsd";
        this.structureXSDFileName = "_0_结构_" + mainXSDFileName;
    }

    public void write() {
        writeXSD(mainXSDFileName, documentMain);
        writeXSD(structureXSDFileName, documentStructure);
    }

    private void writeXSD(String xsdFileName, Document xsdDocument) {
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(new File(
                    xsdFileName));
            OutputFormat format = OutputFormat.createPrettyPrint();
            XMLWriter xmlWriter = new XMLWriter(fileOutputStream, format);
            xmlWriter.write(xsdDocument);
            xmlWriter.close();
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
}
