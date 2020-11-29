package com.jzy.exception;

/**
 * @author JinZhiyun
 * @version 1.0
 * @ClassName InvalidParameterException
 * @description 不合法的入参异常
 * @date 2019/11/19 9:34
 **/
public class InvalidParameterException extends RuntimeException {
    private static final long serialVersionUID = -1534690918130787560L;

    public InvalidParameterException() {
    }

    public InvalidParameterException(String message) {
        super(message);
    }
}
