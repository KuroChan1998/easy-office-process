package com.jzy.office.exception;

/**
 * @ClassName NotInitException
 * @Author JinZhiyun
 * @Description 未初始化异常
 * @Date 2021/1/17 16:43
 * @Version 1.0
 **/
public class NotInitException extends Exception {
    public NotInitException() {
    }

    public NotInitException(String message) {
        super(message);
    }
}
