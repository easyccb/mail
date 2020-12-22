package cn.ccb;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.mail.SimpleMailMessage;
import org.springframework.mail.javamail.JavaMailSenderImpl;
import org.springframework.mail.javamail.MimeMessageHelper;

import javax.mail.MessagingException;
import javax.mail.internet.MimeMessage;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@SpringBootTest
class MailApplicationTests {

    @Autowired
    public JavaMailSenderImpl sender;

    @Test
    void test1() throws InterruptedException {
        ExcelReader reader = ExcelUtil.getReader("C:/Users/hp/Desktop/codename.xlsx");
        List<Map<String,Object>> infos = reader.readAll();


        for (int i = 0; i < infos.size(); i++) {
            SimpleMailMessage message = new SimpleMailMessage();

            String address = (String) infos.get(i).get("address");

            System.out.println("============"+address);

            message.setSubject("通知");
            message.setText("下午开会！");

            message.setFrom("413977414@qq.com");
            message.setTo(address);

            sender.send(message);

            Thread.sleep(1000);
        }

    }
}
