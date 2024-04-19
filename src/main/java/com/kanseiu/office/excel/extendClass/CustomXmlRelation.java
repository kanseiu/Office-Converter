package com.kanseiu.office.excel.extendClass;

import org.apache.poi.ooxml.POIXMLRelation;

public class CustomXmlRelation extends POIXMLRelation {
    public static final String RELATION = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";

    public CustomXmlRelation() {
        // 指定文件类型、默认扩展名和内容类型
        super(
            "application/xml", 
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml", 
            "/customXml{0}.xml"
        );
    }
}
