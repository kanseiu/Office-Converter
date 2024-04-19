package com.kanseiu.office.excel.utils;

import com.kanseiu.office.constant.enums.FileTypeEnum;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.StringUtils;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelCommonUtil {

    /**
     * 检查是否符合xlsx格式
     */
    public static FileTypeEnum checkExcelFileType(String contentType){
        FileTypeEnum fileTypeEnum = FileTypeEnum.getFileTypeByContentType(contentType);
        if(FileTypeEnum.XLSX.equals(fileTypeEnum)) {
           return  fileTypeEnum;
        }
        throw new RuntimeException("仅支持xlsx格式的excel文件");
    }

    /**
     * 获取 XSSFWorkbook 的 sheet 名称列表
     * @param workbook
     * @return
     */
    public static List<String> getExcelSheetNameList(XSSFWorkbook workbook) {
        List<String> sheetNameList = new ArrayList<>();
        int sheetNum = workbook.getNumberOfSheets();
        for(int i = 0; i < sheetNum; i++) {
            sheetNameList.add(workbook.getSheetName(i));
        }
        return sheetNameList;
    }

    /**
     * 复制 PrintArea
     * @param srcWb
     * @param destWb
     * @param srcSheet
     * @param sheetIndex
     */
    public static int[] copyPrintArea(XSSFWorkbook srcWb, XSSFWorkbook destWb, XSSFSheet srcSheet, int sheetIndex) {
        String printArea = srcWb.getPrintArea(sheetIndex);
        if(StringUtils.hasText(printArea)) {
            // 获取可打印区域对应的起止列序号
            int[] colAndRowIndexRange = getPrintAreaColAndRowNum(printArea);
            // 设置可打印区域
            destWb.setPrintArea(0, colAndRowIndexRange[0], colAndRowIndexRange[1], colAndRowIndexRange[2], colAndRowIndexRange[3]);
            return colAndRowIndexRange;
        } else {
            XSSFRow srcRow = srcSheet.getRow(srcSheet.getFirstRowNum());
            if(Objects.nonNull(srcRow)) {
                return new int[]{srcRow.getFirstCellNum(), srcRow.getLastCellNum(), srcSheet.getFirstRowNum(), srcSheet.getLastRowNum()};
            } else {
                return new int[]{0, 0, srcSheet.getFirstRowNum(), srcSheet.getLastRowNum()};
            }
        }
    }

    /**
     * // 获取可打印区域的 列、行号
     * @param printAreaStr  可打印区域字符串
     * @return
     */
    public static int[] getPrintAreaColAndRowNum(String printAreaStr) {
        // 分别是：开始列、结束列、开始行、结束行
        int[] colAndRowNumArray = {0, 0, 0, 0};
        // 例如 履历!$A$1:$K$43， 提取 A、1、K、43
        String regex = "\\$(\\w+)\\$(\\d+):\\$(\\w+)\\$(\\d+)";
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(printAreaStr);
        if(matcher.find()) {
            String startColumn = matcher.group(1);
            String startRow = matcher.group(2);
            String endColumn = matcher.group(3);
            String endRow = matcher.group(4);
            // 使用CellReference将列字母转换为列索引
            int startColumnIndex = CellReference.convertColStringToIndex(startColumn);
            int endColumnIndex = CellReference.convertColStringToIndex(endColumn);
            int startRowNum = Integer.parseInt(startRow) - 1;
            int endRowNum = Integer.parseInt(endRow) - 1;
            colAndRowNumArray[0] = startColumnIndex;
            colAndRowNumArray[1] = endColumnIndex;
            colAndRowNumArray[2] = startRowNum;
            colAndRowNumArray[3] = endRowNum;
        }
        return colAndRowNumArray;
    }

}
