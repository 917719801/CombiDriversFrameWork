package cn.gloryroad.util;


import org.openqa.selenium.By;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

//此工具类主要用于从objcectMap文件中获取定位表达式
public class ObjectMap {
    Properties properties;

    public ObjectMap(String propFile) {
        properties = new Properties();
        try {
            FileInputStream in = new FileInputStream(propFile);
            properties.load(in);
            in.close();
        } catch (IOException e) {
            System.out.println("读取对象文件出错");
            e.printStackTrace();
        }

    }


    public By getLocator(String ElementNameInpropFile) throws Exception {
        String locator= properties.getProperty(ElementNameInpropFile);
        String locatorType = locator.split(">")[0];
        String locatorValue = locator.split(">")[1];
        locatorValue = new String(locatorValue.getBytes("ISO-8859-1"),"UTF-8");
        System.out.println("获取的定位类型："+locatorType+"\t获取的定位表达式"+locatorValue);

        if (locatorType.toLowerCase().equals("id")){
            return By.id(locatorValue);
        }else if (locatorType.toLowerCase().equals("name")){
            return By.name(locatorValue);
        }else if ((locatorType.toLowerCase().equals("classname")) || (locatorType.toLowerCase().equals("class"))){
            return By.className(locatorValue);
        }else if ((locatorType.toLowerCase().equals("tagname")) || (locatorType.toLowerCase().equals("tag"))){
            return By.tagName(locatorValue);
        }else if ((locatorType.toLowerCase().equals("linktext")) || (locatorType.toLowerCase().equals("link"))){
            return By.linkText(locatorValue);
        }else if (locatorType.toLowerCase().equals("partiallinktext")){
            return By.partialLinkText(locatorValue);
        }else if ((locatorType.toLowerCase().equals("cssselector")) || (locatorType.toLowerCase().equals("css"))){
            return By.cssSelector(locatorValue);
        }else if (locatorType.toLowerCase().equals("xpath")){
            return By.xpath(locatorValue);
        }else {
            throw  new Exception("输入的locatortype未在程序中被定义："+locatorType);
        }

    }
}
