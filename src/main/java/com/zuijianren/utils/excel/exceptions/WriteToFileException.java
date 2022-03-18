package com.zuijianren.utils.excel.exceptions;

/**
 * 写入文件失败
 *
 * @author: 姜辞旧
 * @since: 2022/3/16 16:25
 * @version: 1.0
 */
public class WriteToFileException extends RuntimeException {

    public WriteToFileException(String s) {
        super(s);
    }
}
