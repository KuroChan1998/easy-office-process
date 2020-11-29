package com.jzy.word;

import com.jzy.exception.InvalidFileTypeException;
import org.apache.poi.xwpf.usermodel.Document;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

/**
 * @ClassName DefaultWord
 * @Author JinZhiyun
 * @Description 默认word文档处理类。一般普通的处理word文档可以继承此类
 * @Date 2020/11/28 19:20
 * @Version 1.0
 **/
public class DefaultWord extends Word {
    public DefaultWord(String inputFile) throws IOException, InvalidFileTypeException {
        super(inputFile);
    }

    public DefaultWord(File file) throws IOException, InvalidFileTypeException {
        super(file);
    }

    public DefaultWord(InputStream inputStream, WordVersionEnum version) throws IOException, InvalidFileTypeException {
        super(inputStream, version);
    }

    public DefaultWord(Document document) {
        super(document);
    }

    public DefaultWord(WordVersionEnum version) throws InvalidFileTypeException {
        super(version);
    }
}
