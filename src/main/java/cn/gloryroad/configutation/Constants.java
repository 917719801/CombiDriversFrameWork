package cn.gloryroad.configutation;

public class Constants {
    public static String projecctDir;
    public static String packageName;
    public static String packageParentDir;
    //获取包路径
    public static String parentDirPath = getParentpath();
    //Excel文件路径
    public static String CONBIDRIVER_DATA_EXCEL=parentDirPath+"\\data\\混合驱动测试用例及数据.xlsx";
    //获取配置文件路径
    public static String CONBIDRIVER_CONFIG_File=parentDirPath+"\\data\\objectMap.properties";
    //“测试用例集合”Sheet中的列号常量设定
    //测试用例Sheet名
    public static final String TESTCASE_SHEET_NAME="测试用例集合";
    //测试用例名称在第一列，用0表示
    public static final int TESTCASE_TESTCASEID=0;
    //第3列用2表示，为测试是否运行
    public static final int TESTCASE_ISEXECUTE=2;
    //测试结果列
    public static final int TESTCASE_TESTRSULT=3;
    //测试数据列
    public static final int TESTCASE_DATASOURCE_SHEETNAME=4;

    //测试步骤sheet的列号常量设定
    //关键字列
    public static final int TESTSTEP_KEYWORDS=3;
    //定位表达式列
    public static final int TESTSTEP_LOCATORTYPE=4;
    //操作值列
    public static final int TESTSTEP_OPERATEVALUE=5;
    //测试结果列
    public static final int TESTSTEP_TESTRESULT=6;

    //数据驱动sheet中列号常量设定
    //是否执行列
    public static final int DATASOURCE_ISEXECUTE=5;
    //测试执行结果列
    public static final int DATASOURCE_RESUULT=6;

    //获取类所在目录的父目录的绝对路径方法
    private static String getParentpath() {
        projecctDir = System.getProperty("user.dir") + "\\src\\main\\java";
        packageName = Constants.class.getPackage().getName();
        packageParentDir = packageName.substring(0, packageName.lastIndexOf("."));
        packageParentDir = String.format("/%s/", packageParentDir.contains(".") ? packageParentDir.replaceAll("\\.", "/") : packageParentDir);
        System.out.println(packageParentDir);
        System.out.println(projecctDir);
        return projecctDir + packageParentDir;
    }
}
