package cn.gloryroad.configutation;
/*
封装关键字
 */

import cn.gloryroad.testScript.TestSuiteByExcel;
import cn.gloryroad.util.KeyBoardUtil;
import cn.gloryroad.util.Log;
import cn.gloryroad.util.ObjectMap;
import cn.gloryroad.util.WaitUitl;
import org.apache.log4j.xml.DOMConfigurator;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.Assert;

import javax.security.auth.login.LoginContext;

import java.util.List;

import static cn.gloryroad.util.WaitUitl.waitWebElement;

public class KeyWordsAction {
    //声明静态webdriver对象，用于在此类中相关driver进行操作
    public static WebDriver driver;
    //声明存储定位表达配置文件objectmap对象
    private static ObjectMap objectMap = new ObjectMap(Constants.CONBIDRIVER_CONFIG_File);
    static {
        DOMConfigurator.configure("log4j.xml");
    }
    //此方法的名称对应Excel文件“关键字”列中的open_browse关键字
    public static void open_browse(String string,String browserName){
        if (browserName.equals("chrome")) {
            System.setProperty("webdriver.chrome.diver", "D:\\WebDriver\\chromedriver.exe");
            driver = new ChromeDriver();
            Log.info("chrome浏览器实例已经声明");
        } else if (browserName.equals("firefox")) {
            System.setProperty("webdriver.gecko.driver", "D:\\WebDriver\\geckodriver.exe");
            driver = new FirefoxDriver();
            Log.info("firefox浏览器实例已经声明");
        } else {
            System.setProperty("webdriver.edge.driver", "D:\\WebDriver\\msedgedriver.exe");
            driver = new EdgeDriver();
            Log.info("edge浏览器实例已经声明");
        }
    }

    public static void navigate(String string,String url){
        driver.get(url);
        Log.info("浏览器访问网址+"+url);
    }

    public static void switch_frame(String locatorExpression,String string){
        try{
            driver.switchTo().frame(driver.findElement(objectMap.getLocator(locatorExpression)));
            Log.info("进入frame"+locatorExpression+"成功");
        }catch (Exception e){
            TestSuiteByExcel.testResult=false;
            Log.info("进入frame"+locatorExpression+"失败，具体异常信息："+e.getMessage());
            e.printStackTrace();
        }
    }
    public static void default_frame(String string1,String string2){
        try{
            Thread.sleep(3000);
            driver.switchTo().defaultContent();
            Log.info("退出frame成功");
        }catch (Exception e){
            TestSuiteByExcel.testResult=false;
            Log.info("退出frame失败，具体异常信息："+e.getMessage());
            e.printStackTrace();
        }
    }
    public static void input(String locatorExpression,String inputString){
        try{
            driver.findElement(objectMap.getLocator(locatorExpression)).clear();
            Log.info("清除"+locatorExpression+"输入框中的所有内容");
            driver.findElement(objectMap.getLocator(locatorExpression)).sendKeys(inputString);
            Log.info("在"+locatorExpression+"输入框中输入"+inputString);
        }catch (Exception e){
            TestSuiteByExcel.testResult=false;
            Log.info("在"+locatorExpression+"输入框中输入"+inputString+"时出现异常，具体异常信息："+e.getMessage());
            e.printStackTrace();
        }
    }
    public static void click(String locatorExpression, String string){
        try{
            driver.findElement(objectMap.getLocator(locatorExpression)).click();
            Log.info("单击"+locatorExpression+"页面元素成功");
        }catch (Exception e){
            TestSuiteByExcel.testResult=false;
            Log.info("单击"+locatorExpression+"页面元素失败，具体异常信息："+e.getMessage());
            e.printStackTrace();
        }
    }
    public static void click_Sure(String locatorExpression,String sureclick){
        try{
            if (sureclick.equals("是")){
                driver.findElement(objectMap.getLocator(locatorExpression)).click();
                Log.info("单击"+locatorExpression+"页面元素成功");
            }
        }catch (Exception e){
            TestSuiteByExcel.testResult=false;
            Log.info("单击"+locatorExpression+"页面元素失败，具体异常信息："+e.getMessage());
            e.printStackTrace();
        }
    }

    public static void  WaitFor_Element(String locatorExpression,String string){
        try{
            waitWebElement(driver,objectMap.getLocator(locatorExpression));
            Log.info("显示等待元素出现成功，元素是："+locatorExpression);
        }catch (Exception e){
            TestSuiteByExcel.testResult=false;
            Log.info("显示等待元素出现时出现异常，具体异常信息："+locatorExpression);
            e.printStackTrace();
        }
    }
    public static void press_Tab(String string1,String string2){
        try{
            Thread.sleep(2000);
            KeyBoardUtil.pressTabKey();
            Log.info("按Tab键成功");
        }catch (Exception e){
            TestSuiteByExcel.testResult=false;
            Log.info("按Tab键时出现异常，具体异常信息："+e.getMessage());
            e.printStackTrace();
        }
    }

    public static void pasteString(String string,String pasteContent){
        try{
            KeyBoardUtil.setAndctrlVClipboardData(pasteContent);
            Log.info("成功粘贴邮件正文："+pasteContent);
        }catch (Exception e){
            TestSuiteByExcel.testResult=false;
            Log.info("在输入框粘贴内容是出现异常，具体异常信息："+e.getMessage());
            e.printStackTrace();
        }
    }
    public static void press_enter(String string1,String string2){
        try{
            KeyBoardUtil.pressEnterKey();
            Log.info("按Enter键成功");
        }catch (Exception e){
            TestSuiteByExcel.testResult=false;
            Log.info("按Enter键时出现异常，具体异常信息："+e.getMessage());
            e.printStackTrace();
        }
    }

    public static void sleep(String string,String sleepTime){
        try {
            WaitUitl.sleep(Integer.parseInt(sleepTime));
            Log.info("休眠"+Integer.parseInt(sleepTime)/1000+"秒成功");
        }catch (Exception e){
            TestSuiteByExcel.testResult=false;
            Log.info("线程休眠时出现异常，具体异常信息："+e.getMessage());
            e.printStackTrace();
        }
    }
    public static void click_sendMailButton(String locatorExpression,String string){
        try{
            List<WebElement> buttons= driver.findElements(objectMap.getLocator(locatorExpression));
            buttons.get(0).click();
            Log.info("单击发送邮件按钮成功");
        }catch (Exception e){
            TestSuiteByExcel.testResult=false;
            Log.info("单击发送邮件按钮失败，具体异常信息："+e.getMessage());
            e.printStackTrace();
        }
    }
    public static void Assert_String(String string,String assertString){
        try{
            Assert.assertTrue(driver.getPageSource().contains(assertString));
            Log.info("成功断言关键字"+assertString);
        }catch (AssertionError e){
            TestSuiteByExcel.testResult=false;
            Log.info("出现断言失败，具体断言失败信息："+e.getMessage());
            System.out.println("断言失败");
        }
    }


    public static void close_browser(String string1,String string2){
        try{
            System.out.println("浏览器关闭函数被执行" );
            Log.info("关闭浏览器窗口");
            driver.quit();
        }catch (Exception e){
            TestSuiteByExcel.testResult=false;
            Log.info("关闭浏览器出现异常，具体异常信息："+e.getMessage());
            e.printStackTrace();
        }
    }
}
