package com.gannett.mail;

/**
 * Created by dmurugan on 5/9/17.
 * Junit to test GetMail class
 */

import net.minidev.json.JSONObject;
import org.junit.Test;

import java.util.List;

public class TestGetMail {
    @Test
    public void testGetMail() {


        //Date is hardcoded - This is a part of unit testing - change the date,mailID and mail password on unit testing
        List<JSONObject> list = GetMail.readMailFromInbox("2017-05-09",1000,"******@gannett.com","*****","Inbox");
        System.out.println(list);

        assert true;

    }
}
