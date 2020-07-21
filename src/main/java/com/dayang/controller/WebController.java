package com.dayang.controller;

import com.dayang.service.WebSerivce;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Controller;

@Controller
public class WebController {

    protected static final Logger logger = LoggerFactory.getLogger(WebController.class);

    @Autowired
    private WebSerivce webSerivce ;

//    @Scheduled(cron="0/3 * * * * ?")//每天22号早上10点30分自动运行
    public void getProductMessage() {
        logger.info("开始启动");
        webSerivce.downLoadData();
    }
}
