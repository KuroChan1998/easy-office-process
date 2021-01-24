package com.jzy.office;

import lombok.Getter;
import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * @ClassName AbstractOffice
 * @Author JinZhiyun
 * @Description office文件处理基类
 * @Date 2021/1/16 14:44
 * @Version 1.0
 **/
public abstract class AbstractOffice {
    /**
     * 输入文件路径
     */
    @Getter
    protected String inputFilePath;

    /**
     * 输出流
     */
    protected OutputStream os;


    /**
     * 将当前修改保存覆盖至输入文件inputFilePath中，并关闭所有流
     *
     * @throws IOException
     */
    public void saveAndClose() throws IOException {
        save();
        close();
    }


    /**
     * 将当前修改保存到输出流，并关闭所有流
     *
     * @param outputStream 输出流
     * @throws IOException
     */
    public void saveAndClose(OutputStream outputStream) throws IOException {
        save(outputStream);
        close();
    }

    /**
     * 将当前修改保存到outputPath对应的文件中，并关闭所有流
     *
     * @param outputPath 输出文件的路径
     * @throws IOException
     */
    public void saveAndClose(String outputPath) throws IOException {
        save(outputPath);
        close();
    }


    /**
     * 将当前修改保存覆盖至输入文件inputFilePath中
     *
     * @throws IOException
     */
    public void save() throws IOException {
        if (StringUtils.isNotEmpty(inputFilePath)) {
            os = new FileOutputStream(new File(inputFilePath));
            save(os);
        } else {
            throw new IOException("文件的默认路径（源文件路径）不存在");
        }
    }

    /**
     * 将当前修改保存到输出流，子类实现具体写入到什么对象中
     *
     * @param outputStream 输出流
     * @throws IOException
     */
    public abstract void save(OutputStream outputStream) throws IOException;

    /**
     * 将当前修改保存到outputPath对应的文件中
     *
     * @param outputPath 输出文件的路径
     * @throws IOException
     */
    public void save(String outputPath) throws IOException {
        os = new FileOutputStream(new File(outputPath));
        save(os);
    }

    /**
     * 关闭流
     *
     * @throws IOException
     */
    public void close() throws IOException {
        if (os != null) {
            os.close();
        }
    }
}
