package com.kanseiu.office.excel.extendClass;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xssf.usermodel.XSSFPictureData;

// https://stackoverflow.com/questions/67666241/how-to-add-svg-image-to-xssfworkbook
public class CXSSFPictureData extends XSSFPictureData {

    protected CXSSFPictureData() {
        super();
    }

    protected CXSSFPictureData(PackagePart part) {
        super(part);
    }

}