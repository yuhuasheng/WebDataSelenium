package com.dayang.service.impl;

import com.dayang.controller.WebController;
import com.dayang.service.WebSerivce;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;
@Service
public class WebServiceImpl implements WebSerivce {

    protected static final Logger logger = LoggerFactory.getLogger(WebController.class);

    private static WebDriver driver = null;

    @Override
    public void downLoadData() {
        //开启浏览器
        startBrowser();
        //获取菜单栏的Recombinant Proteins标签
        WebElement Recombinant  = getProductMenu();
        openUrlNewBlank(Recombinant);
    }

    /**
     * 开启浏览器
     */
    private void startBrowser() {
//        System.setProperty("webdriver.chrome.driver", "C:\\Users\\Administrator\\Downloads\\chromedriver_win32\\chromedriver.exe");
        System.setProperty("webdriver.chrome.driver","C:\\chromedriver_win32\\chromedriver.exe");
        driver = new ChromeDriver();
        driver.get("https://www.creativebiomart.net/");
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
        driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
    }

    /**
     * 获取产品下拉框的产品标签
     * @return
     */
    private WebElement getProductMenu() {
        WebElement cssmenu = driver.findElement(By.id("cssmenu"));
        if(null == cssmenu) {
            return null;
        }
        WebElement ul = cssmenu.findElement(By.tagName("ul"));
        WebElement li = ul.findElements(By.className("has-sub")).get(0);
        WebElement childul = li.findElement(By.tagName("ul"));
        List<WebElement> liElements = childul.findElements(By.tagName("li"));
        WebElement RecombinantLi = liElements.get(0);
        WebElement aElement = RecombinantLi.findElement(By.tagName("a"));
        return aElement;
    }


    private void openUrlNewBlank(WebElement Recombinant) {
        Recombinant.click();
        driver.findElement(By.cssSelector("body")).sendKeys(Keys.CONTROL +"t");
        ArrayList<String> tabs = new ArrayList<String> (driver.getWindowHandles());
        driver.switchTo().window(tabs.get(1)); //switches to new tab
        driver.get("https://www.baidu.com");
    }
}
