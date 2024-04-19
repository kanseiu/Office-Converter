package com.kanseiu.office.excel.utils;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.CollectionUtils;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelPackagePartUtil {

    /**
     * 递归向workbook注册 PackagePart 关联的内容
     * @param workbook      {}
     * @param packagePart   {}
     */
    public static void recursiveRegistPart(XSSFWorkbook workbook, PackagePart packagePart) throws Exception{
        List<PackagePart> packagePartList = findRelationPackagePart(packagePart);
        if(!CollectionUtils.isEmpty(packagePartList)) {
            for (PackagePart part : packagePartList) {
                if(!workbook.getPackage().containPart(part.getPartName())) {
                    workbook.getPackage().registerPartAndContentType(part);
                }
                recursiveRegistPart(workbook, part);
            }
        }
    }

    public static List<PackagePart> findRelationPackagePart(PackagePart packagePart) throws InvalidFormatException {
        List<PackagePart> relatedPartList = new ArrayList<>();
        if(packagePart.hasRelationships()) {
            for (PackageRelationship relationship : packagePart.getRelationships()) {
                relatedPartList.add(packagePart.getRelatedPart(relationship));
            }
        }
        return relatedPartList;
    }

    public static Map<PackagePart, PackageRelationship> findRelationAndPackagePart(PackagePart packagePart) throws InvalidFormatException {
        Map<PackagePart, PackageRelationship> map = new HashMap<>();
        if(packagePart.hasRelationships()) {
            for (PackageRelationship relationship : packagePart.getRelationships()) {
                map.put(packagePart.getRelatedPart(relationship), relationship);
            }
        }
        return map;
    }

    public static PackagePart findPackagePartByPartName(List<PackagePart> wbAllPackPartList, String partName) {
        for (PackagePart wbPackagePart : wbAllPackPartList) {
            if(partName.equals(wbPackagePart.getPartName().getName())) {
                return wbPackagePart;
            }
        }
        throw new RuntimeException("未在workbook中找到名为[" + partName + "]的PackagePart");
    }

}
