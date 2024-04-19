package com.kanseiu.office.controller.excel;

import com.kanseiu.office.constant.enums.FileTypeEnum;
import com.kanseiu.office.excel.CopySheetToNewWorkbook;
import com.kanseiu.office.excel.utils.ExcelCommonUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.CollectionUtils;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * User: kanseiu
 * Date: 2024/4/18
 * Project: OfficeConverter
 * Package: com.kanseiu.office.controller.excel
 */
@RestController
@RequestMapping("splitExcelSheetToNewWorkbook")
public class SplitExcelSheetToNewWorkbookCl {

    @PostMapping
    public void exec(@RequestPart("file") MultipartFile file, @RequestPart(value = "sheetName", required = false) String sheetName, HttpServletResponse response) {
        if(Objects.nonNull(file) && !file.isEmpty()) {
            FileTypeEnum xlsxFile = ExcelCommonUtil.checkExcelFileType(file.getContentType());

            try (XSSFWorkbook srcWb = new XSSFWorkbook(file.getInputStream())){

                // 设置待处理的sheet页名称列表
                List<String> waitHandleList = StringUtils.hasText(sheetName) ? Collections.singletonList(sheetName) : ExcelCommonUtil.getExcelSheetNameList(srcWb);

                // 复制sheet到新workbook
                Map<String, XSSFWorkbook> xssfWorkbookMap = new HashMap<>();
                for (String waitCopySheetName : waitHandleList) {
                    XSSFWorkbook destWb = CopySheetToNewWorkbook.handle(srcWb, waitCopySheetName);
                    xssfWorkbookMap.put(waitCopySheetName, destWb);
                }

                // 导出
                if(!CollectionUtils.isEmpty(xssfWorkbookMap)) {
                    // 设置 HttpServletResponse
                    String fileName = "excel" + FileTypeEnum.ZIP.suffixWithDot;
                    response.setHeader("Content-Disposition", "attachment; filename=" + fileName);
                    response.setContentType(FileTypeEnum.ZIP.contentType);
                    // 导出为ZIP文件
                    try (ZipOutputStream zipOut = new ZipOutputStream(response.getOutputStream())){

                        for (String workbookName : xssfWorkbookMap.keySet()) {
                            XSSFWorkbook destWb = xssfWorkbookMap.get(workbookName);

                            // 将Workbook写入ByteArrayOutputStream
                            ByteArrayOutputStream bos = new ByteArrayOutputStream();
                            destWb.write(bos);
                            destWb.close();

                            // 添加新的ZIP条目，并将其写入ZipOutputStream
                            ZipEntry zipEntry = new ZipEntry(workbookName + xlsxFile.suffixWithDot);
                            zipOut.putNextEntry(zipEntry);
                            zipOut.write(bos.toByteArray());
                            zipOut.closeEntry();

                            bos.close();
                        }
                    }
                }
            } catch (IOException e) {
                throw new RuntimeException("出现异常", e);
            }
        } else {
            throw new RuntimeException("文件有问题");
        }
    }

}
