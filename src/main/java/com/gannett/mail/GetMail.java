package com.gannett.mail;

import microsoft.exchange.webservices.data.autodiscover.IAutodiscoverRedirectionUrl;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.search.FolderTraversal;
import microsoft.exchange.webservices.data.core.enumeration.service.DeleteMode;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.FolderSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.Attachment;
import microsoft.exchange.webservices.data.property.complex.EmailAddress;
import microsoft.exchange.webservices.data.property.complex.FileAttachment;
import microsoft.exchange.webservices.data.property.complex.FolderId;
import microsoft.exchange.webservices.data.search.FindFoldersResults;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.FolderView;
import microsoft.exchange.webservices.data.search.ItemView;
import net.minidev.json.JSONObject;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.safety.Whitelist;


import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;


/**
 * Created by dmurugan on 4/5/17.
 *
 */
public class GetMail {

    private static Logger logger = Logger.getLogger("GetMail");

    private static class RedirectionUrlCallback implements IAutodiscoverRedirectionUrl {
        /**
         * Private function used to redirect http tp https
         * @param redirectionUrl Url that has to be redirected
         *
         */
        public boolean autodiscoverRedirectionUrlValidationCallback(
                String redirectionUrl) {
            return redirectionUrl.toLowerCase().startsWith("https://");
        }
    }

    /**
     * Private function used to get all the items for the date passed
     * @param date Date in String since it is passed from scala api
     * @param no_of_mails total no of mails to be read
     * @param mailID Mail id that has to be used
     * @param password Mail Password that has to be used
     * @return List<Item>
     *
     */
    private static List<Item> getItems(String date, int no_of_mails, String mailID, String password, String folderName) {
        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
        ExchangeCredentials credentials = new WebCredentials(mailID, password);
        service.setCredentials(credentials);

        FindItemsResults<Item> findResults = new FindItemsResults<Item>();

        try {
            service.autodiscoverUrl(mailID, new RedirectionUrlCallback());


            FolderView folderView = new FolderView(100);

            //Just Initializing with default Value which is Inbox
            FolderId folderId = FolderId.getFolderIdFromWellKnownFolderName(WellKnownFolderName.Inbox);

            folderView.setTraversal(FolderTraversal.Deep);
            FindFoldersResults findFolderResults = service.findFolders(WellKnownFolderName.Root, folderView);
            //find specific folder
            for(Folder f : findFolderResults)
            {
                //Find folderId of the folder folderName
                if (f.getDisplayName() == folderName){
                    folderId = f.getId();
                }

            }


            // Bind to the Inbox.
            Folder inbox = Folder.bind(service, folderId);
            inbox.getPermissions();
            ItemView view = new ItemView(no_of_mails);
            findResults = service.findItems(inbox.getId(),view);
            service.loadPropertiesForItems(findResults, PropertySet.FirstClassProperties);
        }
        catch(Exception e) {
            e.printStackTrace();
            logger.log(Level.SEVERE,e.toString());
        }


        try {

        }
        catch (Exception e) {

        }



        List<Item> itemList = new ArrayList<Item>();

        for (Item item : findResults.getItems()) {
            Calendar cal = Calendar.getInstance();
            cal.add(Calendar.DATE, -1);
            DateFormat df = new SimpleDateFormat("yyyy-MM-dd");
            Date startDate;
            try {
                startDate = df.parse(date);
            }
            catch (Exception e) {
                startDate = cal.getTime();
            }

            try {
                if (DateUtils.isSameDay(item.getDateTimeReceived(), startDate)) {
                    itemList.add(item);
                }
            }
             catch (ServiceLocalException e) {
                e.printStackTrace();
                logger.log(Level.SEVERE,e.toString());
            }


        }

        return itemList;
    }

    /**
     * Public function used to get all the mails in JSONObject for the date passed
     * @param date Date in String since it is passed from scala api
     * @param no_of_mails total no of mails to be read
     * @param mailID Mail id that has to be used
     * @param password Mail Password that has to be used
     * @return List<JSONObject>
     *
     */
    public static List<JSONObject> readMailFromInbox(String date, int no_of_mails, String mailID, String password, String folderName) {

        long st = System.currentTimeMillis();
        logger.info("Start time::"+st);

        List<Item> findResults = getItems(date, no_of_mails, mailID, password, folderName);

        List<JSONObject> jsonList = new ArrayList<JSONObject>();

        for (Item item : findResults) {
           try {
                    JSONObject json = new JSONObject();
                    parseMail(item,json);
                    jsonList.add(json);

            }
            catch (ServiceLocalException e) {
                e.printStackTrace();
                logger.log(Level.SEVERE,e.toString());
            }


        }
        long et = System.currentTimeMillis();
        logger.info("End time::"+et);
        logger.info("Time Taken::"+(et-st)/1000);
        return jsonList;

    }
    /**
     * Public function used to delete all the mails for the date passed
     * @param date Date in String since it is passed from scala api
     * @param no_of_mails total no of mails to be read
     * @param mailID Mail id that has to be used
     * @param password Mail Password that has to be used
     * @return boolean
     *
     */

    public static boolean deleteMailFromInbox(String date, int no_of_mails, String mailID, String password, String folderName) throws Exception {

        List<Item> findResults = getItems(date, no_of_mails, mailID, password, folderName);

        for (Item item : findResults) {

            try {
                item.delete(DeleteMode.MoveToDeletedItems);
            } catch (Exception e) {
                e.printStackTrace();
                logger.log(Level.SEVERE,e.toString());
                return false;
            }


        }
        return true;

    }

    /**
     * private function used parse the mail
     * @param item which is the mail item
     * @param json which is part of call be reference
     *
     */
    private static void parseMail(Item item, JSONObject json) throws ServiceLocalException{

        json.put("subject", item.getSubject());

        String textOnly = br2nl(item.getBody().toString());

        json.put("subject", item.getSubject());
        json.put("body", textOnly);
        json.put("fromName",  item.getLastModifiedName());
        json.put("receivedTime", item.getDateTimeReceived());
        json.put("fromAddress", ((EmailMessage) item).getFrom().getAddress());
        json.put("fromType", ((EmailMessage) item).getFrom().getRoutingType());
        json.put("receivedDate", "" + new SimpleDateFormat("yyyy-MM-dd").format(item.getDateTimeReceived()));
        json.put("toAddress", ((EmailMessage) item).getReceivedBy().getAddress());
        json.put("toRecipients", ((EmailMessage) item).getReceivedBy().getName());

        String ccRecipients = "";
        String ccAddress = "";
        for (EmailAddress address : ((EmailMessage) item).getCcRecipients().getItems()) {
            /**
             * String is compared with "" since the IsEmpty function is being deprecated
             * Refer http://stackoverflow.com/questions/3321526/should-i-use-string-isempty-or-equalsstring
             */
            if (ccAddress.equals("")) {
                ccAddress = address.getAddress();
                ccRecipients = address.getName();
            } else {
                ccAddress = ccAddress + "," + address.getAddress();
                ccRecipients = ccRecipients + "," + address.getName();
            }

        }
        json.put("toCCAddress", ccAddress);
        json.put("toCCRecipients", ccRecipients);


        String bccRecipients = "";
        String bccAddress = "";
        for (EmailAddress address : ((EmailMessage) item).getBccRecipients().getItems()) {
            if (bccAddress.equals("")) {
                bccAddress = address.getAddress();
                bccRecipients = address.getName();
            } else {
                bccAddress = bccAddress + "," + address.getAddress();
                bccRecipients = bccRecipients + "," + address.getName();
            }

        }
        json.put("toBCCAddress", bccAddress);
        json.put("toBCCRecipients", bccRecipients);
        Iterator it = item.getCategories().getIterator();
        String categories = "";
        while(it.hasNext()) {
            categories +=":"+it.next().toString();
        }
        json.put("Categories",categories);
        json.put("Importance", item.getImportance());

        for (Attachment attachment : item.getAttachments()) {
            try {
                attachment.load();
                json.put("attachmentContentType", attachment.getContentType());

                json.put("attachmentName", attachment.getName());
                String extension  = FilenameUtils.getExtension(attachment.getName());
                if (extension.equals("txt") || extension.equals("doc") || extension.equals("docx"))
                    json.put("attachmentContent", new String(((FileAttachment) attachment).getContent(), "UTF-8"));
            }
            catch (Exception e){
                e.printStackTrace();
                logger.log(Level.SEVERE,e.toString());

            }

        }


    }

    /**
     * private function used clean html
     * Copied from http://stackoverflow.com/a/19602313/27938
     * @param html html String which has to be cleaned
     * @return String
     *
     */
    private static String br2nl(String html) {
        if(html==null)
            return null;
        Document document = Jsoup.parse(html);
        document.outputSettings(new Document.OutputSettings().prettyPrint(false));//makes html() preserve linebreaks and spacing
        document.select("br").append("\\n");
        document.select("p").prepend("\\n\\n");
        String s = document.html().replaceAll("\\\\n", "\n");
        return Jsoup.clean(s, "", Whitelist.none(), new Document.OutputSettings().prettyPrint(false));
    }


}
