package com.kanseiu.office.constant.enums;

import lombok.AllArgsConstructor;
import lombok.Getter;

import java.util.Arrays;

/**
 * User: kanseiu
 * Date: 2024/4/18
 * Project: OfficeConverter
 * Package: com.kanseiu.office.constant.enums
 */
@Getter
@AllArgsConstructor
public enum FileTypeEnum {

    XLSX("xlsx", ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "excel"),
    XLS("xls", ".xls", "application/vnd.ms-excel", "excel"),
    ZIP("zip", ".zip", "application/zip", "压缩文件"),
    ;

    public final String suffix;

    public final String suffixWithDot;

    public final String contentType;

    public final String desc;

    public static FileTypeEnum getFileTypeByContentType(String contentType) {
        return Arrays.stream(values()).filter(fileTypeEnum -> fileTypeEnum.getContentType().equals(contentType)).findFirst().orElse(null);
    }
}
