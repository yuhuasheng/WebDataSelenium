package com.dayang.download;

import com.dayang.controller.WebController;
import com.dayang.domain.ProteinsInfo;
import com.dayang.util.Tools;
import com.dayang.util.Utils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.*;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;
import java.util.concurrent.TimeUnit;

/**
 * describe:
 *
 * @author 230257
 * @date 2020/07/17
 */
public class WebProductDownLoad {

    protected static final Logger logger = LoggerFactory.getLogger(WebController.class);

    private static WebDriver driver = null;

    /**
     * 总页数
     */
    private static String count = "";

    /**
     * 当前页码
     */
    private static String currentCount = "";
    /**
     * 下一页按钮
     */
    private static WebElement nextPage = null;

    /**
     * 当前句柄
     */
    private static String currentWindow;
    /**
     * 初始的URL
     */
    private static final String ORIGINURL = "https://www.creativebiomart.net/product/recombinant-proteins_1.htm";

    /**
     * 文件路径
     */
    private static final String FILEPATH = "src/main/resources/file/url.txt";

    /**
     * 文件前缀
     */
    private static String FILEPREFIX = "";

    public static void main(String[] args) {

        /*String filePath = Tools.createFile(FILEPATH);
        if ("".equals(filePath)) {
            return;
        }
        //开启浏览器
        startBrowser(ORIGINURL);
        //将url保存到TXT文本中
        recordUrlList(filePath);*/
        //获取文本内容的url
        getUrlList();
    }


    /**
     * 开启浏览器
     */
    private static void startBrowser(String url) throws AWTException, InterruptedException {
        System.setProperty("webdriver.chrome.driver", "C:\\chromedriver_win32\\chromedriver.exe");
//        System.setProperty("webdriver.chrome.driver" ,"D:\\chromedriver_win32\\chromedriver.exe") ;
        driver = new ChromeDriver();
        driver.get(url);
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
        driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);


        /*Robot r = new Robot();
        r.keyPress(KeyEvent.VK_CONTROL);
        r.keyPress(KeyEvent.VK_T);
        r.keyRelease(KeyEvent.VK_CONTROL);
        r.keyRelease(KeyEvent.VK_T);

        Thread.sleep(2000);
        ArrayList<String> tabs = new ArrayList<String>(driver.getWindowHandles());*/
    }

    @Test
    public void test() throws Exception {
        System.setProperty("webdriver.chrome.driver", "C:\\chromedriver_win32\\chromedriver.exe");
        driver = new ChromeDriver();
        String baseUrl = "http://www.google.co.uk/";
        driver.get(baseUrl);



        /*driver.findElement(By.cssSelector("body")).sendKeys(Keys.CONTROL +"t");
        ArrayList<String> tabss = new ArrayList<String> (driver.getWindowHandles());
        driver.switchTo().window(tabss.get(1)); //switches to new tab
        driver.get("https://www.facebook.com");*/


        /*Thread.sleep(2000);
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
        driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
*/
        //To open a new tab
        Robot r = new Robot();
        r.keyPress(KeyEvent.VK_CONTROL);
        r.keyPress(KeyEvent.VK_T);
        r.keyRelease(KeyEvent.VK_CONTROL);
        r.keyRelease(KeyEvent.VK_T);

        Thread.sleep(2000);
        ArrayList<String> tabs = new ArrayList<String>(driver.getWindowHandles());
        driver.switchTo().window(tabs.get(1)); //switches to new tab

        driver.get("https://www.facebook.com");

    }

    /**
     * 获取从A-Z的url, 存放到url.txt文本中
     *
     * @param filePath 文件的绝对路径
     * @return
     */
    private static void recordUrlList(String filePath) {
        WebElement aDsearchleft = driver.findElement(By.className("ADsearchleft"));
        WebElement pElement = aDsearchleft.findElement(By.tagName("p"));
        List<WebElement> aElements = pElement.findElements(By.tagName("a"));
        if (null == aElements || aElements.size() <= 0) {
            System.err.println("获取字母A-Z链接失败...");
            return;
        }
        String href = "";
        for (WebElement element : aElements) {
            href = element.getAttribute("href");
            System.out.println("==>>href: " + href);
            Tools.write(filePath, href);
        }
    }

    /**
     * 获取文本内容的url
     */
    private static void getUrlList() {
        List<String> resultList = Tools.getTextContent("/file/url.txt");
        for (String url : resultList) {
            List<ProteinsInfo> proteinsInfoList = new ArrayList<>();
            try {
                //连接地址中的字母
                FILEPREFIX = url.substring(url.indexOf("_") + 1, url.lastIndexOf("_"));
                if(url.contains("A")) {
                    url = "https://www.creativebiomart.net/Product/ClassSearch?tag=A&classid=1&page=7";
                }
                System.out.println("==>>FILEPREFIX: " + FILEPREFIX);
                System.out.println("==>>url: " + url);
                //开始处理
                startHandler(url, proteinsInfoList);
                //遍历成功则关闭浏览器,开启新的窗口
                driver.close();
            } catch (Exception e) {
                e.printStackTrace();
                //输出至Excel中
                exportExcelFile(proteinsInfoList);
                System.err.println("循环遍历失败...");
                return;
            }
            //输出至Excel中
            exportExcelFile(proteinsInfoList);
        }

    }


    /**
     * 开始处理
     *
     * @param url 网页路径
     */
    private static void startHandler(String url, List<ProteinsInfo> proteinsInfoList) throws Exception {
        //打开网页
        startBrowser(url);
        //获取产品列表
        getProductTable(proteinsInfoList);
    }


    /**
     * 获取产品列表
     *
     * @throws Exception
     */
    private static void getProductTable(List<ProteinsInfo> proteinsInfoList) throws Exception {
        //获取下一页和下一页按钮
        getNextTag();
        System.out.println("==>>count: " + count);
        System.out.println("==>>nextPage: " + nextPage);
        System.out.println("==>>currentCount: " + currentCount);
        //需要循环的页数
        int sum = Integer.parseInt(count) - Integer.parseInt(currentCount);
        for (int i = 0; i <= sum; i++) {
            //遍历tr标签
            traversalTrElement(proteinsInfoList);
            //点击下一页
            nextPage.click();
            driver = driver.switchTo().window(driver.getWindowHandle());
            driver.manage().window().maximize();
            driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
            driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
            currentCount = "";
            count = "";
            nextPage = null;
            //获取下一页,总页数和下一页按钮
            getNextTag();
        }

    }

    /**
     * 获取下一页,总页数和下一页按钮
     */
    private static void getNextTag() {
        WebElement scott = driver.findElement(By.className("scott"));
        List<WebElement> aListElements = scott.findElements(By.tagName("a"));
        //当前页码
        currentCount = scott.findElement(By.className("current")).getText();
        //总页码
        count = aListElements.get(aListElements.size() - 2).getText();
        //下一页标签
        nextPage = aListElements.get(aListElements.size() - 1);
    }


    /**
     * 遍历tr标签
     *
     * @throws AWTException
     * @throws InterruptedException
     */
    private static void traversalTrElement(List<ProteinsInfo> proteinsInfoList) throws Exception {
        WebElement tableElement = driver.findElement(By.id("table-breakpoint"));
        WebElement tbody = tableElement.findElement(By.tagName("tbody"));
        List<WebElement> trElements = tbody.findElements(By.tagName("tr"));
        if (null == trElements || trElements.size() <= 0) {
            System.out.println("产品列表为空...");
            return;
        }
        for (WebElement element : trElements) {
            ProteinsInfo proteinsInfo = new ProteinsInfo();
            WebElement tdElement = element.findElements(By.tagName("td")).get(1);
            WebElement span = tdElement.findElement(By.tagName("span"));
            WebElement aElement = span.findElement(By.tagName("a"));
            String href = aElement.getAttribute("href");
            //当前句柄
            currentWindow = driver.getWindowHandle();

            Thread.sleep(2000);
            Robot robot = new Robot();
            robot.keyPress(KeyEvent.VK_CONTROL);
            robot.keyPress(KeyEvent.VK_T);
            robot.keyRelease(KeyEvent.VK_CONTROL);
            robot.keyRelease(KeyEvent.VK_T);

            //休眠五秒,防止无法拿到所有的句柄
            Thread.sleep(3000);
            //详情页
            detailPage(href, proteinsInfo, proteinsInfoList);

        }
    }


    /**
     * 获取详情页信息
     */
    private static void detailPage(String url, ProteinsInfo proteinsInfo, List<ProteinsInfo> proteinsInfoList) throws Exception {
        //get all windows
        Set<String> handles = driver.getWindowHandles();
        if(handles.size() <2) {
            throw new Exception("打开新的标签页失败...");
        }
        WebDriver window = null;
        for (String s : handles) {
            if (s.equals(currentWindow)) {
                continue;
            } else {
                window = driver.switchTo().window(s);
                window.get(url);
                window.manage().window().maximize();
                window.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
                window.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
                //获取参数页的标签
                getParamsPage(window, proteinsInfo, proteinsInfoList);

                //close the table window
                window.close();
            }
            driver.switchTo().window(currentWindow);
        }
    }


    /**
     * 获取参数页的标签
     *
     * @param window 当前页面的驱动类
     */
    private static void getParamsPage(WebDriver window, ProteinsInfo proteinsInfo, List<ProteinsInfo> proteinsInfoList) throws Exception {
        String url = window.getCurrentUrl();
        System.out.println("当前页的url为: " + url);
        proteinsInfo.setUrl(url);
        WebElement tab = window.findElement(By.id("tab"));
        WebElement divElement = tab.findElement(By.className("tab-nav"));
        //标签集合
        List<WebElement> aElements = divElement.findElements(By.tagName("a"));
        try {
            for (WebElement element : aElements) {
                String text = element.getText();
                if ("Specification".equals(text)) {
                    clickItem(element);
                    //获取item列表
                    getItem(window, proteinsInfo, proteinsInfoList, "Specification");
                } else if ("Gene Information".equals(text)) {
                    clickItem(element);
                    //获取item列表
                    getItem(window, proteinsInfo, proteinsInfoList, "Gene Information");
                }
//                int i = 1/0;

            }
        } catch (Exception e) {
            e.printStackTrace();
            throw new Exception("产品参数获取失败...");
        } finally {
            if (!proteinsInfoList.contains(proteinsInfo)) {
                proteinsInfoList.add(proteinsInfo);
            }
        }

    }

    /**
     * 点击item
     *
     * @param element 元素
     */
    private static void clickItem(WebElement element) {
        //代表此选项标题不是选中的状态
        if (!"current".equals(element.getAttribute("class"))) {
            element.click();
        }
    }


    /**
     * 获取item列表(Specification/Gene Information)
     *
     * @param window
     * @param proteinsInfo pojo类
     */
    private static void getItem(WebDriver window, ProteinsInfo proteinsInfo, List<ProteinsInfo> proteinsInfoList, String type) {
        WebElement tab = window.findElement(By.id("tab"));
        WebElement tabConElement = tab.findElement(By.className("tab-con"));
        WebElement jTabConElement = tabConElement.findElement(By.className("j-tab-con"));
        WebElement tabConItemElement = null;
        if ("Specification".equals(type)) {
            tabConItemElement = jTabConElement.findElements(By.className("tab-con-item")).get(0);
        } else if ("Gene Information".equals(type)) {
            tabConItemElement = jTabConElement.findElements(By.className("tab-con-item")).get(1);
        }
        WebElement tableElement = tabConItemElement.findElement(By.tagName("table"));
        WebElement tbodyElement = tableElement.findElement(By.tagName("tbody"));
        List<WebElement> trElements = tbodyElement.findElements(By.tagName("tr"));
        for (WebElement trElement : trElements) {
            List<WebElement> tdElements = trElement.findElements(By.tagName("td"));
            //参数名
            String paramsName = tdElements.get(0).getText() == null ? "" : tdElements.get(0).getText();
            paramsName = paramsName.replace(":", "").trim();
            //内容
            String value = tdElements.get(1).getText() == null ? "" : tdElements.get(1).getText();

            //设置参数
            setParams(paramsName, value, proteinsInfo, proteinsInfoList);
        }
    }


    /**
     * 设置参数
     *
     * @param paramsName   列名
     * @param value        值
     * @param proteinsInfo pojo类
     */
    private static void setParams(String paramsName, String value, ProteinsInfo proteinsInfo, List<ProteinsInfo> proteinsInfoList) {
        if ("cat.no.".equals(paramsName.toLowerCase()) || paramsName.toLowerCase().startsWith("cat")) {
            proteinsInfo.setCat(value);
        } else if ("product name".equals(paramsName.toLowerCase())) {
            proteinsInfo.setProductName(value);
        } else if ("product overview".equals(paramsName.toLowerCase())) {
            proteinsInfo.setProductOverview(value);
        } else if ("description".equals(paramsName.toLowerCase())) {
            proteinsInfo.setDescription(value);
        } else if ("source".equals(paramsName.toLowerCase())) {
            proteinsInfo.setSourceHost(value);
        } else if ("species".equals(paramsName.toLowerCase())) {
            proteinsInfo.setSpecies(value);
        } else if ("applications".equals(paramsName.toLowerCase())) {
            proteinsInfo.setApplications(value);
        } else if ("tag".equals(paramsName.toLowerCase())) {
            proteinsInfo.setTag(value);
        } else if ("form".equals(paramsName.toLowerCase())) {
            proteinsInfo.setForm(value);
        } else if ("activity".equals(paramsName.toLowerCase())) {
            proteinsInfo.setActivity(value);
        } else if ("formulation".equals(paramsName.toLowerCase())) {
            proteinsInfo.setFormulation(value);
        } else if ("molecular mass".equals(paramsName.toLowerCase())) {
            proteinsInfo.setMolecularMass(value);
        } else if ("molecular weight".equals(paramsName.toLowerCase())) {
            proteinsInfo.setMolecularWeight(value);
        } else if ("purity".equals(paramsName.toLowerCase())) {
            proteinsInfo.setPurity(value);
        } else if ("concentration".equals(paramsName.toLowerCase())) {
            proteinsInfo.setConcentration(value);
        } else if ("endotoxin".equals(paramsName.toLowerCase())) {
            proteinsInfo.setEndotoxin(value);
        } else if ("predicted n-terminus".equals(paramsName.toLowerCase())) {
            proteinsInfo.setPredictedNTerminus(value);
        } else if ("bio-activity".equals(paramsName.toLowerCase())) {
            proteinsInfo.setBioActivity(value);
        } else if ("unit definition".equals(paramsName.toLowerCase())) {
            proteinsInfo.setUnitDefinition(value);
        } else if ("aa sequence".equals(paramsName.toLowerCase())) {
            proteinsInfo.setAaSequence(value);
        } else if ("protein length".equals(paramsName.toLowerCase())) {
            proteinsInfo.setProteinLength(value);
        } else if ("storage".equals(paramsName.toLowerCase())) {
            proteinsInfo.setStorage(value);
        } else if ("reconstitution".equals(paramsName.toLowerCase())) {
            proteinsInfo.setReconstitution(value);
        } else if ("storage buffer".equals(paramsName.toLowerCase())) {
            proteinsInfo.setStorageBuffer(value);
        } else if ("tissue specificity".equals(paramsName.toLowerCase())) {
            proteinsInfo.setTissueSpecificity(value);
        } else if ("shipping condition".equals(paramsName.toLowerCase())) {
            proteinsInfo.setShippingCondition(value);
        } else if ("stability".equals(paramsName.toLowerCase())) {
            proteinsInfo.setStability(value);
        } else if ("usage".equals(paramsName.toLowerCase())) {
            proteinsInfo.setUsage(value);
        } else if ("quality control test".equals(paramsName.toLowerCase())) {
            proteinsInfo.setQualityControlTest(value);
        } else if ("preservative".equals(paramsName.toLowerCase())) {
            proteinsInfo.setPreservative(value);
        } else if ("sequence similarities".equals(paramsName.toLowerCase())) {
            proteinsInfo.setSequenceSimilarities(value);
        } else if ("gene name".equals(paramsName.toLowerCase())) {
            proteinsInfo.setGeneName(value);
        } else if ("official symbol".equals(paramsName.toLowerCase())) {
            proteinsInfo.setOfficialSymbol(value);
        } else if ("synonyms".equals(paramsName.toLowerCase())) {
            proteinsInfo.setSynonyms(value);
        } else if ("gene id".equals(paramsName.toLowerCase())) {
            proteinsInfo.setGeneId(value);
        } else if ("mrna refseq".equals(paramsName.toLowerCase())) {
            proteinsInfo.setmRNARefseq(value);
        } else if ("protein refseq".equals(paramsName.toLowerCase())) {
            proteinsInfo.setProteinRefseq(value);
        } else if ("mim".equals(paramsName.toLowerCase())) {
            proteinsInfo.setMIM(value);
        } else if ("uniprot id".equals(paramsName.toLowerCase())) {
            proteinsInfo.setUniProtId(value);
        } else if ("chromosome location".equals(paramsName.toLowerCase())) {
            proteinsInfo.setChromosomeLocation(value);
        } else if ("function".equals(paramsName.toLowerCase())) {
            proteinsInfo.setFunction(value);
        }

    }


    /**
     * 输出产品信息
     *
     * @param proteinsInfoList
     */
    private static void exportExcelFile(List<ProteinsInfo> proteinsInfoList) {
        OutputStream stream = null;
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy年MM月dd日  HH时mm分ss秒");
        Date date = new Date();
        try {
            String dateTime = sdf.format(date);
            String filePath = "D:\\网页爬虫\\Recombinant-Proteins\\" + FILEPREFIX + "\\";
            File dir = new File(filePath);
            if (!dir.exists()) {
                dir.mkdirs();
            }
            String fileName = filePath + "\\" + "Recombinant-Proteins_" + FILEPREFIX + "_" + dateTime + ".xls";
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("Recombinant-Proteins");

            //设置列样式
            HSSFCellStyle columnStyle = Utils.setColumnStyle(workbook);

            //设置内容的样式
            HSSFCellStyle contentStyle = Utils.setContentStyle(workbook);

            //设置每一列的宽度
            Utils.setColumnWidth(sheet);

            //设置列名和样式
            Utils.addTitleData(sheet, columnStyle);

            //冻结表头
            Utils.freezeHeader(sheet);

            for (int i = 0; i < proteinsInfoList.size(); i++) {
                ProteinsInfo proteinsInfo = proteinsInfoList.get(i);
                int k = i + 1;
                Utils.createLine(sheet, contentStyle, k, proteinsInfo);
            }
            File file = new File(fileName);
            stream = new FileOutputStream(file);
            workbook.write(stream);
            System.out.println("导出完成！");
            System.out.println("文件路径为: " + file.getAbsolutePath());
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("导出失败！");
        } finally {
            if (stream != null) {
                try {
                    stream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}
