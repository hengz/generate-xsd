package com.vickz.xsdgen;

import java.io.File;

/**
 * This program convert Microsoft Excel file to structure XSD file
 * @author zhangheng
 */
public class XSDGenMain {

    public void generateXSD(String xsdPath) {
        XSDGen xsdGen = new XSDGen();
        xsdGen.parse(new File(xsdPath)).write();

    }

    public static void main(String args[]) {
        XSDGenMain xsdGenMain = new XSDGenMain();
        xsdGenMain.generateXSD("刑事一审.xlsx");
    }
}
