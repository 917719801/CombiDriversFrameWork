package cn.gloryroad.testScript;

import cn.gloryroad.configutation.Constants;
import cn.gloryroad.configutation.KeyWordsAction;
import cn.gloryroad.util.ExcelUtil;
import cn.gloryroad.util.Log;
import org.apache.log4j.xml.DOMConfigurator;
import org.testng.Assert;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import java.lang.reflect.Method;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class TestSuiteByExcel {
    public static Method method[];
    public static String keyword;
    public static String locatorExpression;
    public static String value;
    public static KeyWordsAction keyWordsAction;
    public static int testStep;
    public static int testLastStep;
    public static String testCaseID;
    public static String testCaseRunFlag;
    public static boolean testResult;
    public String testData;
    public static String testDataRunFlag;


    @Test
    public void testTestSuite() {
        //声明关键字动作类的实例
        keyWordsAction = new KeyWordsAction();
        //使用java的反射机制获取keyWordsAction类的所有方法对象
        method = keyWordsAction.getClass().getMethods();
        //定义Excel文件的路径
        String excelFilePath = Constants.CONBIDRIVER_DATA_EXCEL;
        //读取Excel文件中的发送邮件Sheet为操作目标
        ExcelUtil.setExcelFile(excelFilePath);
        //读取测试用例集合sheet中的测试用例总数
        int testCasesCount = ExcelUtil.getRowCount(Constants.TESTCASE_SHEET_NAME);
        //使用for循环执行“测试用例集合”sheet中所有标记为y的测试用例
        for (int testCaseNo = 1; testCaseNo <= testCasesCount; testCaseNo++) {
            //读取“测试用例集合”sheet中每一行的测试用例序号
            testCaseID = ExcelUtil.getCellData(Constants.TESTCASE_SHEET_NAME, testCaseNo, Constants.TESTCASE_TESTCASEID);
            //读取“测试用例集合”每行的是否运行列中的值
            testCaseRunFlag = ExcelUtil.getCellData(Constants.TESTCASE_SHEET_NAME, testCaseNo, Constants.TESTCASE_ISEXECUTE);
            //判断是否运行列中的值是否是Y，是y则执行测试用例中的所有步骤
            if (testCaseRunFlag.trim().equals("y")) {
                Log.startTEstCase(testCaseID);
                //设定测试用例当前结果为true，表名测试执行成功
                testResult = true;
                //获取当前要执行测试用例的第一个步骤所在行的行号
                testStep = ExcelUtil.getFirstRowContainsTestCaseId(testCaseID, testCaseID, Constants.TESTCASE_TESTCASEID);
                //获取当前要执行测试用例最后一步坐在的行号
                testLastStep = ExcelUtil.getTestCaseLastStepRow(testCaseID, testCaseID, testStep);
                //读取"测试用例集合“Sheet中测试数据列的内容
                testData = ExcelUtil.getCellData(Constants.TESTCASE_SHEET_NAME, testCaseNo, Constants.TESTCASE_DATASOURCE_SHEETNAME);
                //如果测试数据不为空，则运行数据驱动
                if (testData.trim().equals("无")) {
                    Log.info("执行关键字驱动用例开始，sheet表名为：" + testData);
                    //遍历测试步骤sheet中的所有测试步骤
                    for (; testStep < testLastStep; testStep++) {
                        //获取测试步骤sheet中读取关键字和操作值
                        keyword = ExcelUtil.getCellData(testCaseID, testStep, Constants.TESTSTEP_KEYWORDS);
                        Log.info("从Excel文件读取到的关键字是：" + keyword);
                        //获取测试步骤sheet中读取的定位表达式
                        locatorExpression = ExcelUtil.getCellData(testCaseID, testStep, Constants.TESTSTEP_LOCATORTYPE);
                        //获取测试步骤sheet中读取的操作值
                        value = ExcelUtil.getCellData(testCaseID, testStep, Constants.TESTSTEP_OPERATEVALUE);
                        Log.info("从Excel文件中读取的操作值是：" + value);
                        executeActions(testCaseID);
                        //判断是否执行成功，并写入Excel文件
                        if (testResult == false) {
                            ExcelUtil.setCellData(Constants.TESTCASE_SHEET_NAME, testCaseNo, Constants.TESTCASE_TESTRSULT, "测试执行失败");
                            Log.endTestCase(testCaseID);
                            break;
                        }
                        if (testResult == true) {
                            ExcelUtil.setCellData(Constants.TESTCASE_SHEET_NAME, testCaseNo, Constants.TESTCASE_TESTRSULT, "测试执行成功");
                        }
                    }
                } else {
                    //读取测试数据sheet中的数据总数
                    int testDatasCount = ExcelUtil.getRowCount(testData);
                    //使用for循环执行所有标记为Y的测试数据
                    for (int testDataNo = 1; testDataNo <= testDatasCount; testDataNo++) {
                        //读取测试数据sheet中每行是否运行列的值
                        testDataRunFlag = ExcelUtil.getCellData(testData, testDataNo, Constants.DATASOURCE_ISEXECUTE);
                        //如果“是否运行“列中的值为”y“，则执行该测试数据
                        if (testDataRunFlag.trim().equals("y")) {
                            Log.info("开始数据驱动用例执行");
                            //遍历测试用例中的所有测试步骤
                            for (; testStep < testLastStep; testStep++) {
                                keyword = ExcelUtil.getCellData(testCaseID, testStep, Constants.TESTSTEP_KEYWORDS);
                                Log.info("从Excel文件读取的关键字是：" + keyword);
                                locatorExpression = ExcelUtil.getCellData(testCaseID, testStep, Constants.TESTSTEP_LOCATORTYPE);
                                value = ExcelUtil.getCellData(testCaseID, testStep, Constants.TESTSTEP_OPERATEVALUE);
                                value = value.trim();
                                Log.info("从Excel文件读取的操作值是：" + value);
                                //若value值为“A-Z”，则替换为测试数据sheet中的数据
                                if (null != value && value.length() == 1 && TestSuiteByExcel.judgeContainsStr(value)) {
//                                    通过Excel文件中的列字母得到的列索引
                                    int conlumnNo = ExcelUtil.excelColStrToNum(value);
                                    value = ExcelUtil.getCellData(testData, testDataNo, conlumnNo);
                                    Log.info("Excel文件中列值为：" + value);
                                }
                                //执行用例
                                executeActions(testCaseID);
                                if (testResult == true) {
                                    ExcelUtil.setCellData(testData, testDataNo, Constants.DATASOURCE_RESUULT, "测试数据执行成功");
                                } else {
                                    ExcelUtil.setCellData(testData, testDataNo, Constants.DATASOURCE_RESUULT, "测试数据执行失败");
                                    KeyWordsAction.close_browser("", "");
                                    break;
                                }
                            }
                            //在测试步骤sheet中获取当前要执行测试步骤第一行的行号
                            testStep = ExcelUtil.getFirstRowContainsTestCaseId(testCaseID, testCaseID, Constants.TESTCASE_TESTCASEID);
                            if (testResult == false) {
                                ExcelUtil.setCellData(Constants.TESTCASE_SHEET_NAME, testCaseNo, Constants.TESTCASE_TESTRSULT, "测试执行失败");
                                Log.endTestCase(testCaseID);
                                break;
                            }
                            if (testResult == true) {
                                ExcelUtil.setCellData(Constants.TESTCASE_SHEET_NAME, testCaseNo, Constants.TESTCASE_TESTRSULT, "测试执行成功");
                            }
                        }
                    }
                }
            }
        }
    }
    //判断是否在a`z之间
    private static boolean judgeContainsStr(String str) {
        String regex = ".*[A-Z]+.*";
        Matcher matcher = Pattern.compile(regex).matcher(str);
        return matcher.matches();
    }

    private static void executeActions(String testCaseID) {
        try {
            for (int i = 0; i < method.length; i++) {
                if (method[i].getName().equals(keyword)) {
                    method[i].invoke(keyWordsAction, locatorExpression, value);
                    if (testResult == true) {
                        ExcelUtil.setCellData(testCaseID, testStep, Constants.TESTSTEP_TESTRESULT, "测试步骤执行成功");
                    } else {
                        ExcelUtil.setCellData(testCaseID, testStep, Constants.TESTSTEP_TESTRESULT, "测试步骤执行失败");
                        KeyWordsAction.close_browser("", "");
                        break;
                    }
                }

            }
        } catch (Exception e) {
            Assert.fail("执行出现异常，测试用例执行失败");
        }
    }


    @BeforeClass
    public void BeforeClass() {
        DOMConfigurator.configure("log4j.xml");
    }

}
