package com.kanseiu.office.excel.utils;

import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.*;
import org.springframework.util.StringUtils;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class ExcelCTUtil {

    public static void copySharedStrings(XSSFWorkbook srcWb, XSSFWorkbook destWb, XSSFSheet destSheet){
        // 复制 sheet.xml 关联的 sharedStrings.xml 内容
        // 获取目标sheet的xml中的<sheetData>标签内容
        CTSheetData destCTSheetData = destSheet.getCTWorksheet().getSheetData();
        // 获取 srcWb 的 xl/sharedStrings.xml 的内容
        SharedStringsTable srcSharedStringsTable = srcWb.getSharedStringSource();
        if(srcSharedStringsTable.getCount() > 0) {
            // 获取 destSheet 的 sheet.xml 中的 <sheetData><row/></sheetData> 内容
            List<CTRow> destCtRowList = destCTSheetData.getRowList();
            // 获取 destWb 的 xl/sharedStrings.xml 的内容
            SharedStringsTable destSharedStringsTable = destWb.getSharedStringSource();
            Map<Integer, Integer> sharedStringIndexMap = new HashMap<>();
            for (CTRow destCtRow : destCtRowList) {
                List<CTCell> destCtCellList = destCtRow.getCList();
                for (CTCell destCtCell : destCtCellList) {
                    String sharedStringsTableIndexStr = destCtCell.getV();
                    // 过滤掉没有值的单元格和类型不是字符串的
                    // 当单元格类型不是字符串的时候，比如数字，getV() 得到的值，就不是 sharedStrings.xml 中的值，而是代表实际的数字，这个是合理的，不能为每个数字建立 sharedStrings.xml
                    // STCellType.S 表示共享字符串，即该单元格内的字符串，存储在 sharedStrings 中
                    if(StringUtils.hasText(sharedStringsTableIndexStr) && destCtCell.getT().equals(STCellType.S)) {
                        // 将 待复制sheet的 sharedStrings 复制到新 workbook 的 sharedStrings.xml，并重置sheet.xml的v值
                        int sharedStringsTableIndex = Integer.parseInt(sharedStringsTableIndexStr);
                        RichTextString richTextString = srcSharedStringsTable.getItemAt(sharedStringsTableIndex);
                        // 由于单元格引用的 SharedString 序号，可能重复，因此不能反复添加，已经添加过的，直接获取
                        int newIndex;
                        if(sharedStringIndexMap.containsKey(sharedStringsTableIndex)) {
                            newIndex = sharedStringIndexMap.get(sharedStringsTableIndex);
                        } else {
                            newIndex = destSharedStringsTable.addSharedStringItem(richTextString);
                            sharedStringIndexMap.put(sharedStringsTableIndex, newIndex);
                        }
                        destCtCell.setV(String.valueOf(newIndex));
                    }
                }
            }
        }
    }

    /**
     * 从 sheet.xml 提取样式引用ID，为复制 styles.xml 的内容做准备
     */
    public static List<Long> extractStyleRef(XSSFSheet srcSheet) {
        // 复制sheet页的所有列的默认样式索引
        List<Long> srcSheetCellStyleRefIdList = new ArrayList<>();
        for (CTCols srcCtCols : srcSheet.getCTWorksheet().getColsList()) {
            for (CTCol srcCtCol : srcCtCols.getColList()) {
                srcSheetCellStyleRefIdList.add(srcCtCol.getStyle());
            }
        }

        // 复制sheet页的所有单元格的自定义样式索引
        // 复制sheet页的所有行的默认样式索引
        for (CTRow srcCtRow : srcSheet.getCTWorksheet().getSheetData().getRowList()) {
            // 复制行默认样式索引
            srcSheetCellStyleRefIdList.add(srcCtRow.getS());
            // 复制单元格自定义样式索引
            for (CTCell srcCtCell : srcCtRow.getCList()) {
                srcSheetCellStyleRefIdList.add(srcCtCell.getS());
            }
        }
        // 去重
        return srcSheetCellStyleRefIdList.stream().distinct().collect(Collectors.toList());
    }

    /**
     * 复制 style.xml 中 font、fill、border、numFmts和单元格样式，并修改相应的引用ID
     */
    public static void copyAndSetStyle(XSSFWorkbook srcWb, XSSFWorkbook destWb, XSSFSheet destSheet, List<Long> srcSheetCellStyleRefIdList) {

        StylesTable destStylesTable = destWb.getStylesSource();
        StylesTable srcStylesTable = srcWb.getStylesSource();

        // 复制 style.xml 中的 font
        int destWbInitFontSize = destStylesTable.getFonts().size();
        for(int i = 0; i < srcStylesTable.getFonts().size(); i++) {
            XSSFFont srcFont = srcStylesTable.getFontAt(i);
            // 覆盖destWb默认的font
            if(destWbInitFontSize > 0) {
                destStylesTable.getFonts().get(i).getCTFont().set(srcFont.getCTFont());
                destWbInitFontSize--;
            } else {
                destStylesTable.putFont(srcFont, true);
            }
        }

        // 复制 style.xml 中的 fill
        int destWbInitFillSize = destStylesTable.getFills().size();
        for(int i = 0; i < srcStylesTable.getFills().size(); i++) {
            XSSFCellFill srcFill = srcStylesTable.getFillAt(i);
            // 覆盖destWb默认的fill
            if(destWbInitFillSize > 0) {
                destStylesTable.getFills().get(i).getCTFill().set(srcFill.getCTFill());
                destWbInitFillSize--;
            } else {
                destStylesTable.putFill(srcFill);
            }
        }

        // 复制 style.xml 中的 border
        int destWbBorderSize = destStylesTable.getBorders().size();
        for(int i = 0; i < srcStylesTable.getBorders().size(); i++) {
            XSSFCellBorder srcBorder = srcStylesTable.getBorderAt(i);
            // 覆盖destWb默认的border
            if(destWbBorderSize > 0) {
                destStylesTable.getBorders().get(i).getCTBorder().set(srcBorder.getCTBorder());
                destWbBorderSize--;
            } else {
                destStylesTable.putBorder(srcBorder);
            }
        }

        // 复制 style.xml 中的 cellStyleXf
        int destWbCellStyleXfSize = destStylesTable._getStyleXfsSize();
        for(int i = 0; i < srcStylesTable._getStyleXfsSize(); i++) {
            CTXf srcCTXf = srcStylesTable.getCellStyleXfAt(i);
            // 覆盖destWb默认的cellStyleXf
            if(destWbCellStyleXfSize > 0) {
                destStylesTable.getCellStyleXfAt(i).set(srcCTXf);
                destWbCellStyleXfSize--;
            } else {
                destStylesTable.putCellStyleXf(srcCTXf);
            }
        }

        // 复制 style.xml 中的 numFmts
        Map<Short, String> srcNumFmtMap = srcStylesTable.getNumberFormats();
        srcNumFmtMap.forEach(destStylesTable::putNumberFormat);

        // 复制 cellXf
        Map<Long, Long> cellStyleRefIdMap = new HashMap<>();
        int cellStyleRefIndex = 0;
        int destWbCellXfSize = destStylesTable.getNumCellStyles();
        for (Long refId : srcSheetCellStyleRefIdList) {
            // getCellXfAt from 0
            CTXf srcCellXf = srcStylesTable.getCellXfAt(refId.intValue());
            // 覆盖destWb默认的cellCTXf
            if(destWbCellXfSize > 0) {
                destStylesTable.getCellXfAt(cellStyleRefIndex).set(srcCellXf);
                destWbCellXfSize--;
            } else {
                destStylesTable.putCellXf(srcCellXf);
            }
            cellStyleRefIdMap.put(refId, (long) cellStyleRefIndex);
            cellStyleRefIndex++;
        }

        // 修改 dest sheet.xml 中的 style.xml cellXfs 引用
        // 修改列默认样式
        for (CTCols destCtCols : destSheet.getCTWorksheet().getColsList()) {
            for (CTCol srcCtCol : destCtCols.getColList()) {
                long orginStyleRef = srcCtCol.getStyle();
                long nowStyleRef = cellStyleRefIdMap.get(orginStyleRef);
                srcCtCol.setStyle(nowStyleRef);
            }
        }

        // 修改所有单元格的自定义样式引用的索引
        // 修改行默认样式索引
        for (CTRow destCtRow : destSheet.getCTWorksheet().getSheetData().getRowList()) {
            // 修改行默认样式索引
            long orginRowStyleRef = destCtRow.getS();
            long nowRowStyleRef = cellStyleRefIdMap.get(orginRowStyleRef);
            destCtRow.setS(nowRowStyleRef);
            // 修改所有单元格的自定义样式引用的索引
            for (CTCell destCtCell : destCtRow.getCList()) {
                long orginCellStyleRef = destCtCell.getS();
                long nowCellStyleRef = cellStyleRefIdMap.get(orginCellStyleRef);
                destCtCell.setS(nowCellStyleRef);
            }
        }
    }

}
