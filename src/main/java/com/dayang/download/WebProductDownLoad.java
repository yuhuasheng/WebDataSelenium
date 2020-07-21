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
    private static void startBrowser(String url) {
        System.setProperty("webdriver.chrome.driver", "C:\\chromedriver_win32\\chromedriver.exe");
        driver = new ChromeDriver();
        driver.get(url);
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
        driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
    }

    @Test
    public void test() throws Exception {
        System.setProperty("webdriver.chrome.driver", "C:\\chromedriver_win32\\chromedriver.exe");
        driver = new ChromeDriver();
        String baseUrl = "http://www.google.co.uk/";
        driver.get(baseUrl);
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
        List<ProteinsInfo> proteinsInfoList = new ArrayList<>();
        for (String url : resultList) {
            try {
                System.out.println("==>>url: " + url);
                startHandler(url, proteinsInfoList);
            } catch (Exception e) {
                e.printStackTrace();
                //输出至Excel中
                exportExcelFile(proteinsInfoList);
                System.err.println("循环遍历失败...");
                break;
            }
        }

        //输出至Excel中
        exportExcelFile(proteinsInfoList);
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
        }

    }

    /**
     * 获取下一页和下一页按钮
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
    private static void traversalTrElement(List<ProteinsInfo> proteinsInfoList) throws AWTException, InterruptedException {
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
            Robot robot = new Robot();
            robot.keyPress(KeyEvent.VK_CONTROL);
            robot.keyPress(KeyEvent.VK_T);
            robot.keyRelease(KeyEvent.VK_CONTROL);
            robot.keyRelease(KeyEvent.VK_T);

            //休眠五秒,防止无法拿到所有的句柄
            Thread.sleep(5000);
            //详情页
            detailPage(href, proteinsInfo);

            //添加内容
            proteinsInfoList.add(proteinsInfo);
        }
    }



    /**
     * 获取详情页信息
     */
    private static void detailPage(String url, ProteinsInfo proteinsInfo) {
        //get all windows
        Set<String> handles = driver.getWindowHandles();
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
                getParamsPage(window, proteinsInfo);

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
    private static void getParamsPage(WebDriver window, ProteinsInfo proteinsInfo) {
        String url = window.getCurrentUrl();
        System.out.println("当前页的url为: " + url);
        proteinsInfo.setUrl(url);
        WebElement tab = window.findElement(By.id("tab"));
        WebElement divElement = tab.findElement(By.className("tab-nav"));
        //标签集合
        List<WebElement> aElements = divElement.findElements(By.tagName("a"));
        for (WebElement element : aElements) {
            String text = element.getText();
            if ("Specification".equals(text)) {
                clickItem(element);
                //获取item列表
                getItem(tab, proteinsInfo);
            } else if ("Gene Information".equals(text)) {
                clickItem(element);
                //获取item列表
                getItem(tab, proteinsInfo);
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
     * @param tab
     * @param proteinsInfo pojo类
     */
    private static void getItem(WebElement tab, ProteinsInfo proteinsInfo) {
        WebElement tabConElement = tab.findElement(By.className("tab-con"));
        WebElement jTabConElement = tabConElement.findElement(By.className("j-tab-con"));
        WebElement tabConItemElement = jTabConElement.findElements(By.className("tab-con-item")).get(0);
        WebElement tableElement = tabConItemElement.findElement(By.tagName("table"));
        WebElement tbodyElement = tableElement.findElement(By.tagName("tbody"));
        List<WebElement> trElements = tbodyElement.findElements(By.tagName("tr"));
        for (WebElement trElement : trElements) {
            List<WebElement> tdElements = trElement.findElements(By.tagName("td"));
            //参数名
            String paramsName = tdElements.get(0).getText();
            //内容
            String value = tdElements.get(1).getText() == null ? "" : tdElements.get(1).getText();
            //设置参数
            setParams(paramsName, value, proteinsInfo);
        }
    }


    /**
     * 设置参数
     *
     * @param paramsName   列名
     * @param value        值
     * @param proteinsInfo pojo类
     */
    private static void setParams(String paramsName, String value, ProteinsInfo proteinsInfo) {
        if (paramsName.toLowerCase().startsWith("cat")) {
            proteinsInfo.setCat(value);
        } else if (paramsName.toLowerCase().startsWith("product name")) {
            proteinsInfo.setProductName(value);
        } else if (paramsName.toLowerCase().startsWith("product overview")) {
            proteinsInfo.setProductOverview(value);
        } else if (paramsName.toLowerCase().startsWith("description")) {
            proteinsInfo.setDescription(value);
        } else if (paramsName.toLowerCase().startsWith("source")) {
            proteinsInfo.setSourceHost(value);
        } else if (paramsName.toLowerCase().startsWith("species")) {
            proteinsInfo.setSpecies(value);
        } else if (paramsName.toLowerCase().startsWith("applications")) {
            proteinsInfo.setApplications(value);
        } else if (paramsName.toLowerCase().startsWith("tag")) {
            proteinsInfo.setTag(value);
        } else if (paramsName.toLowerCase().startsWith("form")) {
            proteinsInfo.setForm(value);
        } else if (paramsName.toLowerCase().startsWith("activity")) {
            proteinsInfo.setActivity(value);
        } else if (paramsName.toLowerCase().startsWith("formulation")) {
            proteinsInfo.setFormulation(value);
        } else if (paramsName.toLowerCase().startsWith("molecular mass")) {
            proteinsInfo.setMolecularMass(value);
        } else if (paramsName.toLowerCase().startsWith("molecular weight")) {
            proteinsInfo.setMolecularWeight(value);
        } else if (paramsName.toLowerCase().startsWith("purity")) {
            proteinsInfo.setPurity(value);
        } else if (paramsName.toLowerCase().startsWith("concentration")) {
            proteinsInfo.setConcentration(value);
        } else if (paramsName.toLowerCase().startsWith("endotoxin")) {
            proteinsInfo.setEndotoxin(value);
        } else if (paramsName.toLowerCase().startsWith("predicted n-terminus")) {
            proteinsInfo.setPredictedNTerminus(value);
        } else if (paramsName.toLowerCase().startsWith("bio-activity")) {
            proteinsInfo.setBioActivity(value);
        } else if (paramsName.toLowerCase().startsWith("unit definition")) {
            proteinsInfo.setUnitDefinition(value);
        } else if (paramsName.toLowerCase().startsWith("aa sequence")) {
            proteinsInfo.setAaSequence(value);
        } else if (paramsName.toLowerCase().startsWith("protein length")) {
            proteinsInfo.setProteinLength(value);
        } else if (paramsName.toLowerCase().startsWith("storage")) {
            proteinsInfo.setStorage(value);
        } else if (paramsName.toLowerCase().startsWith("reconstitution")) {
            proteinsInfo.setReconstitution(value);
        } else if (paramsName.toLowerCase().startsWith("storage buffer")) {
            proteinsInfo.setStorageBuffer(value);
        } else if (paramsName.toLowerCase().startsWith("tissue specificity")) {
            proteinsInfo.setTissueSpecificity(value);
        } else if (paramsName.toLowerCase().startsWith("shipping condition")) {
            proteinsInfo.setShippingCondition(value);
        } else if (paramsName.toLowerCase().startsWith("stability")) {
            proteinsInfo.setStability(value);
        } else if (paramsName.toLowerCase().startsWith("usage")) {
            proteinsInfo.setUsage(value);
        } else if (paramsName.toLowerCase().startsWith("quality control test")) {
            proteinsInfo.setQualityControlTest(value);
        } else if (paramsName.toLowerCase().startsWith("preservative")) {
            proteinsInfo.setPreservative(value);
        } else if (paramsName.toLowerCase().startsWith("sequence similarities")) {
            proteinsInfo.setSequenceSimilarities(value);
        } else if (paramsName.toLowerCase().startsWith("gene name")) {
            proteinsInfo.setGeneName(value);
        } else if (paramsName.toLowerCase().startsWith("official symbol")) {
            proteinsInfo.setOfficialSymbol(value);
        } else if (paramsName.toLowerCase().startsWith("synonyms")) {
            proteinsInfo.setSynonyms(value);
        } else if (paramsName.toLowerCase().startsWith("gene id")) {
            proteinsInfo.setGeneId(value);
        } else if (paramsName.toLowerCase().startsWith("mrna refseq")) {
            proteinsInfo.setmRNARefseq(value);
        } else if (paramsName.toLowerCase().startsWith("protein refseq")) {
            proteinsInfo.setProteinRefseq(value);
        } else if (paramsName.toLowerCase().startsWith("mim")) {
            proteinsInfo.setMIM(value);
        } else if (paramsName.toLowerCase().startsWith("uniprot id")) {
            proteinsInfo.setUniProtId(value);
        } else if (paramsName.toLowerCase().startsWith("chromosome location")) {
            proteinsInfo.setChromosomeLocation(value);
        } else if (paramsName.toLowerCase().startsWith("function")) {
            proteinsInfo.setFunction(value);
        }
    }


    /**
     * 输出产品信息
     * @param proteinsInfoList
     */
    private static void exportExcelFile(List<ProteinsInfo> proteinsInfoList) {
        OutputStream stream = null;
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy年MM月dd日  HH时mm分ss秒");
        Date date = new Date();
        try {
            String dateTime = sdf.format(date);
            String filePath = "D:\\网页爬虫\\Recombinant-Proteins\\";
            File dir = new File(filePath);
            if (!dir.exists()) {
                dir.mkdirs();
            }
            String fileName = filePath + "\\" + "Recombinant-Proteins_" + dateTime + ".xls";
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
