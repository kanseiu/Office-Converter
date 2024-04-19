package com.kanseiu.office.excel;

import com.kanseiu.office.constant.Common;
import com.kanseiu.office.excel.utils.ExcelCTUtil;
import com.kanseiu.office.excel.utils.ExcelCommonUtil;
import com.kanseiu.office.excel.utils.ExcelPackagePartUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheet;

import java.io.*;
import java.util.List;
import java.util.Map;
import java.util.Objects;

@Slf4j
public class CopySheetToNewWorkbook {

    public static XSSFWorkbook handle(XSSFWorkbook srcWb, String waitCopySheetName) {
        try {
            // 获取需要复制的 sheet 的 CTSheet
            CTSheet srcCtSheet = srcWb.getCTWorkbook().getSheets().getSheetList().stream().filter(r -> waitCopySheetName.equals(r.getName())).findFirst().orElseThrow(() -> new RuntimeException("未找到[" + waitCopySheetName + "]"));

            // 获取 srcSheet 和 sheetIndex
            XSSFSheet srcSheet = srcWb.getSheet(waitCopySheetName);
            int sheetIndex = srcWb.getSheetIndex(srcSheet);

            // 获取整个workbook的PackagePart
            List<PackagePart> srcWbAllPackPartList = srcWb.getPackagePart().getPackage().getParts();

            // 找到 workbook 的 PackagePart
            PackagePart srcWorkbookPackagePart = ExcelPackagePartUtil.findPackagePartByPartName(srcWbAllPackPartList, Common.WORKBOOK_PART_NAME);

            // 从 workbook 的 PackagePart 找到 待复制的 worksheet 的 PackagePart
            PackageRelationship waitCopySheetPackageRelationship = srcWorkbookPackagePart.getRelationship(srcCtSheet.getId());
            if(Objects.isNull(waitCopySheetPackageRelationship)) {
                throw new RuntimeException("未在workbook中找到[" + waitCopySheetName + "]的 RelationShip");
            }
            PackagePart waitCopySheetPackagePart = srcWorkbookPackagePart.getRelatedPart(waitCopySheetPackageRelationship);
            if(Objects.isNull(waitCopySheetPackagePart)) {
                throw new RuntimeException("未在workbook中找到[" + waitCopySheetName + "]的 PackagePart");
            }

            // 创建新的workbook
            XSSFWorkbook destWb = new XSSFWorkbook();

            // 将 waitCopySheetPackagePart 关联的内容，复制过来
            ExcelPackagePartUtil.recursiveRegistPart(destWb, waitCopySheetPackagePart);

            // 创建新的 worksheet
            XSSFSheet destSheet = destWb.createSheet(waitCopySheetName);

            PackagePart destSheetPackagePart = destSheet.getPackagePart();

            // 将 srcSheet 的 relation 复制给 destSheet
            Map<PackagePart, PackageRelationship> relationAndPackagePart = ExcelPackagePartUtil.findRelationAndPackagePart(waitCopySheetPackagePart);
            relationAndPackagePart.forEach(((packagePart, packageRelationship) -> {
                destSheetPackagePart.addRelationship(
                        packagePart.getPartName(),
                        packageRelationship.getTargetMode(),
                        packageRelationship.getRelationshipType(),
                        packageRelationship.getId());
            }));

            // 复制 src sheet.xml 的内容
            destSheet.getCTWorksheet().set(srcSheet.getCTWorksheet());

            // 复制 sheet.xml 关联的 sharedStrings.xml 内容
            ExcelCTUtil.copySharedStrings(srcWb, destWb, destSheet);

            //  复制 printArea, 返回sheet页面的起止colnum
            int[] sheetColIndexRange = ExcelCommonUtil.copyPrintArea(srcWb, destWb, srcSheet, sheetIndex);

            // 从 sheet.xml 提取样式引用ID，为复制 styles.xml 的内容做准备
            List<Long> srcSheetCellStyleRefIdList = ExcelCTUtil.extractStyleRef(srcSheet);

            // 复制 style.xml 整体的样式结构 和 命名空间
            destWb.getStylesSource().getCTStylesheet().set(srcWb.getStylesSource().getCTStylesheet());

            // 复制 style.xml 中 font、fill、border、numFmts和单元格样式，并修改相应的引用ID
            ExcelCTUtil.copyAndSetStyle(srcWb, destWb, destSheet, srcSheetCellStyleRefIdList);

            // 返回
            return destWb;
        } catch (IOException e) {
            throw new RuntimeException("发生IO错误", e);
        } catch (InvalidFormatException e) {
            throw new RuntimeException("findRelationAndPackagePart发生错误", e);
        } catch (Exception e) {
            throw new RuntimeException("发生未知错误", e);
        }
    }
}
