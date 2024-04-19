package com.kanseiu.office.excel.extendClass;

import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFFactory;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.List;

// https://stackoverflow.com/questions/67666241/how-to-add-svg-image-to-xssfworkbook
public class CXSSFWorkbook extends XSSFWorkbook {

    public int addSVGPicture(InputStream is) throws Exception {
        Field _xssfFactory = XSSFWorkbook.class.getDeclaredField("xssfFactory");
        _xssfFactory.setAccessible(true);
        XSSFFactory xssfFactory = (XSSFFactory) _xssfFactory.get(this);

        int imageNumber = getAllPictures().size() + 1;

        Field _pictures = XSSFWorkbook.class.getDeclaredField("pictures");
        _pictures.setAccessible(true);

        @SuppressWarnings("unchecked")
        List<XSSFPictureData> pictures = (List<XSSFPictureData>) _pictures.get(this);

        CXSSFPictureData img = createRelationship(CXSSFRelation.IMAGE_SVG, xssfFactory, imageNumber, true).getDocumentPart(); // see MyXSSFPictureData.java and MyXSSFRelation.java
        try (OutputStream out = img.getPackagePart().getOutputStream()) {
            IOUtils.copy(is, out);
        }
        pictures.add(img);

        return imageNumber - 1;
    }

}