package com.kanseiu.office.excel.extendClass;

import org.apache.poi.ooxml.POIXMLRelation;
import org.apache.poi.openxml4j.opc.PackageRelationshipTypes;

// https://stackoverflow.com/questions/67666241/how-to-add-svg-image-to-xssfworkbook
public final class CXSSFRelation extends POIXMLRelation {

    public static final CXSSFRelation IMAGE_SVG = new CXSSFRelation(
            "image/svg",
            PackageRelationshipTypes.IMAGE_PART,
            "/xl/media/image#.svg",
            CXSSFPictureData::new, CXSSFPictureData::new
    );

    private CXSSFRelation(String type, String rel, String defaultName,
                          NoArgConstructor noArgConstructor,
                          PackagePartConstructor packagePartConstructor) {
        super(type, rel, defaultName, noArgConstructor, packagePartConstructor, null);
    }

}