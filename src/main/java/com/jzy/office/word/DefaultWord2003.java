package com.jzy.office.word;

import com.jzy.office.exception.InvalidFileTypeException;
import lombok.ToString;
import org.apache.poi.hwpf.HWPFDocument;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

/**
 * @ClassName DefaultWord2003
 * @Author JinZhiyun
 * @Description 默认word 2003文档处理类。一般普通的处理word 2003文档可以继承此类
 * @Date 2021/1/23 20:15
 * @Version 1.0
 **/
@ToString(callSuper = true)
public class DefaultWord2003 extends Word2003 {
    private static final long serialVersionUID = -20032L;

    public DefaultWord2003(String inputFile) throws IOException, InvalidFileTypeException {
        super(inputFile);
    }

    public DefaultWord2003(File file) throws IOException, InvalidFileTypeException {
        super(file);
    }

    public DefaultWord2003(InputStream inputStream, WordVersionEnum version) throws IOException, InvalidFileTypeException {
        super(inputStream, version);
    }

    public DefaultWord2003(HWPFDocument document) {
        super(document);
    }
}
