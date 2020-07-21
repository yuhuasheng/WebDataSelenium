package com.dayang;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.scheduling.annotation.EnableScheduling;

@SpringBootApplication
@EnableScheduling
public class Application {

    protected static final Logger logger = LoggerFactory.getLogger(Application.class) ;

    public static void main(String[] args) {
        logger.info("SpringBoot开始加载") ;
        SpringApplication.run(Application.class) ;

    }
}
