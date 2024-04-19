package com.kanseiu.office.utils;

import com.kanseiu.office.constant.Common;
import org.springframework.util.StringUtils;

public class CustomizeStringUtil {

    public static String getMediaDirPath(String mediaPath){
        if(StringUtils.hasText(mediaPath)) {
            StringBuilder mediaDirPathStringBuilder = new StringBuilder();
            String[] splitStr = mediaPath.split(Common.SLASH);
            for(int i = 0; i < splitStr.length - 1; i++) {
                mediaDirPathStringBuilder.append(Common.SLASH);
                mediaDirPathStringBuilder.append(splitStr[i]);
            }
            return mediaDirPathStringBuilder.toString();
        }
        return null;
    }
}
