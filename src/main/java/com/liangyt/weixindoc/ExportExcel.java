package com.glodon.geb.excel;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.date.DateTime;
import cn.hutool.core.date.DateUtil;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.ReUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.http.ContentType;
import cn.hutool.http.HttpRequest;
import cn.hutool.http.HttpUtil;
import cn.hutool.json.JSONArray;
import cn.hutool.json.JSONObject;

import java.io.File;
import java.io.UnsupportedEncodingException;
import java.net.URLDecoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CompletableFuture;

/**
 * 导出企业微信在线问的那个中的excel
 * TODO 登录cookie处理
 * @author liangyt
 * @date 2023-03-14 9:37
 */

public class ExportExcel {

    public static final String HOST = "https://doc.weixin.qq.com";

    public static final String EXPORT_OFFICE = "/v1/export/export_office";

    public static final String QUERY_PROGRESS = "/v1/export/query_progress";

    public static final String WEB_DISK_LIST = "/webdisk/list";
    //下载路径
    public static final String downloadPath = "C:\\Users\\liangyt\\Desktop\\修改公式\\最新模板\\%s-%s\\";
    //访问cookie
    public static final String cookie = "tdoc_uid=13102702692956305; wedoc_openid=wozbKqDgAAViLbnFW84hrjTo1kptMuvg; wedoc_sid=1pE1TYx4cUIunGNnALM5VQAA; wedoc_sids=13102702692956305&1pE1TYx4cUIunGNnALM5VQAA; wedoc_skey=13102702692956305&df63e231297d3663f308f7a49607bf83; wedoc_ticket=13102702692956305&CAESIH2QeVp8UL-yhRKokpo6WMEp5yJY7Rcxp_4slUmhNOn6";

    public static void main(String[] args) {
        Map<String, String> docIdMap = webDiskList("i.1970325140031928.1688850478091595", "i.1970325140031928.1688850478091595_d.671171005NXaR", "");

        DateTime date = DateUtil.date();
        int day = date.dayOfMonth();
        int month = date.month();
        String path = String.format(downloadPath, month, day);

        if(!FileUtil.exist(path)){
            FileUtil.mkdir(path);
        } else {
            FileUtil.clean(path);
        }

        downloadFile(docIdMap, path);
    }

    protected static void downloadFile(Map<String, String> docIdMap, String path) {
        List<CompletableFuture<Void>> futureList = new ArrayList<>();
        for (String docId : docIdMap.keySet()) {
            CompletableFuture<Void> future = CompletableFuture.runAsync(() -> {
                HttpRequest request = HttpUtil.createPost(HOST + EXPORT_OFFICE);
                request.contentType(ContentType.FORM_URLENCODED.getValue());
                request.cookie(cookie);
                request.form("version", "2");
                request.form("docId", docId);
                String body = request.execute().body();
                String operationId = new JSONObject(body).getStr("operationId");

                HttpRequest queryProcessRequest = HttpUtil.createGet(HOST + QUERY_PROGRESS);
                queryProcessRequest.contentType(ContentType.JSON.getValue());
                queryProcessRequest.cookie(cookie);
                String file_url;
                while (true) {
                    try {
                        Thread.sleep(2000);
                    } catch (InterruptedException e) {
                        e.printStackTrace();
                    }
                    queryProcessRequest.form("operationId", operationId);
                    body = queryProcessRequest.execute().body();
                    JSONObject jsonObject = new JSONObject(body);
                    String status = jsonObject.getStr("status");
                    if("Done".equals(status)){
                        file_url = jsonObject.getStr("file_url");
                        break;
                    }
                }
                String docPath = docIdMap.get(docId);
                if(StrUtil.isNotBlank(docPath) && !docPath.endsWith(File.separator)){
                    docPath += File.separator;
                }

                String fileName = null;
                try {
                    fileName = getFileNameFromDisposition(file_url);
                } catch (UnsupportedEncodingException e) {
                    e.printStackTrace();
                }
                File filePath = new File(path + docPath + fileName);
                System.out.println("开始下载：" + file_url);
                HttpUtil.downloadFile(file_url, filePath);
                System.out.println("下载完成：" + docPath + fileName);
            });
            futureList.add(future);
        }
        futureList.stream().forEach(CompletableFuture::join);
    }

    public static Map<String, String> webDiskList(String spaceId, String fatherId, String parentPath){
        if(StrUtil.isNotBlank(parentPath) && !parentPath.endsWith(File.separator)){
            parentPath += File.separator;
        }
        Map<String, String> docIdMap = new HashMap<>();
        HttpRequest request = HttpUtil.createPost(HOST + WEB_DISK_LIST);
        request.contentType(ContentType.JSON.getValue());
        request.cookie(cookie);
        request.form("func", 11);
        request.form("father_id", fatherId);
        request.form("start", 0);
        request.form("start", 0);
        request.form("limit", 150);
        request.form("space_id", spaceId);
        request.form("sort_type", 1);
        String body = request.execute().body();
        JSONObject jsonObject = new JSONObject(body);
        JSONArray array = (JSONArray)jsonObject.getByPath("body.file_list");
        if(CollUtil.isEmpty(array)){
            return null;
        }
        for (Object o : array) {
            JSONObject jo = (JSONObject)o;
            if("1".equals(jo.getStr("file_type"))){
                String file_id = jo.getStr("file_id");
                try {
                    Thread.sleep(500);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
                Map<String, String> childDocIdMap = webDiskList(spaceId, file_id, parentPath + jo.getStr("name"));
                if(childDocIdMap != null){
                    docIdMap.putAll(childDocIdMap);
                }
            } else {
                docIdMap.put(jo.getStr("doc_id"), parentPath);
            }
        }
        return docIdMap;
    }

    private static String getFileNameFromDisposition(String fileUrl) throws UnsupportedEncodingException {
        fileUrl = URLDecoder.decode(fileUrl, "UTF-8");
        String fileName = ReUtil.get("filename\\*=\"(.*?)\"", fileUrl, 1);
        if (StrUtil.isBlank(fileName)) {
            fileName = StrUtil.subBetween(fileUrl, "filename*=", "&q-sign-algorithm=");
            fileName = fileName.replaceAll("UTF-8''", "");
            fileName = URLDecoder.decode(fileName, "UTF-8");
        }

        return fileName;
    }

}
