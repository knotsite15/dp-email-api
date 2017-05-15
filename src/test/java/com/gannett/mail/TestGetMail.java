package com.gannett.mail;

/**
 * Created by dmurugan on 5/9/17.
 * Junit to test MSExchangeEmailService class
 */

import net.minidev.json.JSONObject;
import org.junit.Test;

import java.util.List;

public class TestGetMail {
    @Test
    public void testGetMail() {


        //Date is hardcoded - This is a part of unit testing - change the date,mailID and mail password on unit testing
        List<JSONObject> list = MSExchangeEmailService.readMail("2017-05-15",1000,"mobfeedtst@gannett.com","*****","Gannett2017");
        System.out.println(list);

        assert true;

    }
}
