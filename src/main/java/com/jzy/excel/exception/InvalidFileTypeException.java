package com.jzy.excel.exception;

/**
 * @author JinZhiyun
 * @version 1.0
 * @ClassName InvalidFileTypeException
 * @description 输入文件类型不符合规则的异常
 * @date 2019/10/30 13:01
 **/
public class InvalidFileTypeException extends Exception {
    private static final long serialVersionUID = 5972322191747094090L;

    public InvalidFileTypeException() {
    }

    public InvalidFileTypeException(String message) {
        super(message);
    }
}
