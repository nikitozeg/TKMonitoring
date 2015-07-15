import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.HttpClientBuilder;
import org.openqa.jetty.util.URI;

import javax.swing.*;
import java.awt.*;
import java.io.*;

import static javax.swing.JFrame.EXIT_ON_CLOSE;

public class MonitoringDL extends JPanel {
    static private final String newline = "\n";
    JButton openButton;
    JTextArea log, input;
    JFileChooser fc;

    Workbook w;
    Double insuranceResponse;
    Double intercity;
    String kladrFrom;
    String kladrTo;
    Double summa;
    Double priceFrom;
    Double priceTO;
    String terminalFrom;
    String terminalTO;
    String insuranceResponseVOZ, intercityVOZ, longitude, latitude, coords, summaVOZ, priceFromVOZ, priceTOVOZ, summaVOZAction;
    String addressLine;
    Double weight, volume, insurance;
    int count;
    int enteredNumber = 0;
    JProgressBar progressBar;
    String path;

    public void setCoords(String address) throws Exception {

        HttpClient httpClient1 = HttpClientBuilder.create().build();
        HttpGet request1 = new HttpGet("https://geocode-maps.yandex.ru/1.x/?results=1&format=json&geocode=" + URI.encodePath(address));

        HttpResponse response1 = httpClient1.execute(request1);
        HttpEntity entity1 = response1.getEntity();
        InputStream instream1 = entity1.getContent();
        BufferedReader reader1 = new BufferedReader(new InputStreamReader(instream1,"UTF-8"));


        StringBuilder sb1 = new StringBuilder();

        String line1 = null;
        try {
            while ((line1 = reader1.readLine()) != null) {
                sb1.append(line1 + "\n");
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                instream1.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        String ss1 = sb1.toString();
        // now you have the string representation of the HTML request
        //   System.out.println("RESPONSE: " + sb1);

        JsonParser parser = new JsonParser();//response.toString()
        JsonObject mainObject = parser.parse(sb1.toString()).getAsJsonObject().getAsJsonObject("response");
        addressLine=mainObject.getAsJsonObject("GeoObjectCollection").getAsJsonArray("featureMember").get(0).getAsJsonObject().getAsJsonObject("GeoObject").getAsJsonObject("metaDataProperty")
                    .getAsJsonObject("GeocoderMetaData").getAsJsonObject("AddressDetails").getAsJsonObject("Country").get("AddressLine").getAsString();
        coords = mainObject.getAsJsonObject("GeoObjectCollection").getAsJsonArray("featureMember").get(0).getAsJsonObject().getAsJsonObject("GeoObject").getAsJsonObject("Point").get("pos").getAsString();
        latitude = coords.substring(coords.indexOf(" ") + 1, coords.length());
        longitude = coords.substring(0, coords.indexOf(" "));
    }

    private String getKladrAndSetCoords(String address) throws Exception {
        setCoords(address);
        String kladr = "";
        HttpClient httpClient1 = HttpClientBuilder.create().build();
        HttpPost request1 = new HttpPost("https://suggestions.dadata.ru/suggestions/api/4_1/rs/suggest/address");
        StringEntity params1 = new StringEntity("{\"count\":1,\"query\":\"" + addressLine + "\"}", "utf-8");
        request1.addHeader("content-type", "application/json");
        request1.addHeader("Authorization", "Token 84beb76a98914195f374779f2f313d31efca3c5d");
        request1.addHeader("X-Secret", "cb82deee2d367b967ba569b5fc11b9e21a8c4832");
        request1.setEntity(params1);

        HttpResponse response1 = httpClient1.execute(request1);
        HttpEntity entity1 = response1.getEntity();
        InputStream instream1 = entity1.getContent();
        BufferedReader reader1 = new BufferedReader(new InputStreamReader(instream1));


        StringBuilder sb1 = new StringBuilder();

        String line1 = null;
        try {
            while ((line1 = reader1.readLine()) != null) {
                sb1.append(line1 + "\n");
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                instream1.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        String ss1 = sb1.toString();
        // now you have the string representation of the HTML request
        //  System.out.println("RESPONSE: " + ss1);

        JsonParser parser = new JsonParser();//response.toString()

        JsonArray mainObject = parser.parse(sb1.toString()).getAsJsonObject().getAsJsonArray("suggestions");
        //System.out.println(mainObject.get(0).getAsJsonObject().getAsJsonObject("data").get("kladr_id").getAsString());
        kladr = mainObject.get(0).getAsJsonObject().getAsJsonObject("data").get("kladr_id").getAsString();

        return kladr;

    }

    public void sendGet() throws Exception {
       // File crowlerResult = new File("output.xls");
          File crowlerResult = new File("C:\\Users\\n.ivanov\\Dropbox\\AutoMonitoringDL\\output.xls");
        path="C:\\Users\\n.ivanov\\Dropbox\\AutoMonitoringDL\\inpu.xls";
        File exlFile = new File(path);
        w = Workbook.getWorkbook(exlFile);
        Sheet sheet = w.getSheet(0);
        WritableWorkbook writableWorkbook = Workbook.createWorkbook(crowlerResult);
        WritableSheet writableSheet = writableWorkbook.createSheet("Sheet2", 0);

        jxl.write.Label label01 = new jxl.write.Label(0, 0, "От");
        jxl.write.Label label02 = new jxl.write.Label(1, 0, "До");
        jxl.write.Label label03 = new jxl.write.Label(2, 0, "Вес");
        jxl.write.Label label04 = new jxl.write.Label(3, 0, "Объем");

        //ДЛ
        jxl.write.Label label05 = new jxl.write.Label(5, 0, "Забор");
        jxl.write.Label label06 = new jxl.write.Label(6, 0, "МТ");
        jxl.write.Label label07 = new jxl.write.Label(7, 0, "Отвоз");
        jxl.write.Label label08 = new jxl.write.Label(8, 0, "Страховка");
        jxl.write.Label label09 = new jxl.write.Label(9, 0, "ИТОГО");
        jxl.write.Label label010= new jxl.write.Label(10, 0, "Т.отправителя");
        jxl.write.Label label011 = new jxl.write.Label(11, 0, "Т.получателя");

        //ВОЗОВОЗ
        jxl.write.Label label13 = new jxl.write.Label(13, 0, "Забор");
        jxl.write.Label label14 = new jxl.write.Label(14, 0, "МТ");
        jxl.write.Label label15 = new jxl.write.Label(15, 0, "Отвоз");
        jxl.write.Label label16 = new jxl.write.Label(16, 0, "Страховка");
        jxl.write.Label label17 = new jxl.write.Label(17, 0, "ИТОГО без скидки");
        jxl.write.Label label18 = new jxl.write.Label(18, 0, "ИТОГО со скидкой");

        writableSheet.addCell(label01);
        writableSheet.addCell(label02);
        writableSheet.addCell(label03);
        writableSheet.addCell(label04);
        writableSheet.addCell(label05);
        writableSheet.addCell(label06);
        writableSheet.addCell(label07);
        writableSheet.addCell(label08);
        writableSheet.addCell(label09);
        writableSheet.addCell(label010);
        writableSheet.addCell(label011);
        writableSheet.addCell(label13);
        writableSheet.addCell(label14);
        writableSheet.addCell(label15);
        writableSheet.addCell(label16);
        writableSheet.addCell(label17);
        writableSheet.addCell(label18);

        HttpClient httpClient = HttpClientBuilder.create().build();
        JsonParser parser = new JsonParser();

        try {

            String to = "";
            String from = "";

           /* BufferedReader br = new BufferedReader(new InputStreamReader(System.in));
            System.out.print("Введите количество обрабатываемых строк:");
            try {
                enteredNumber = Integer.parseInt(br.readLine());
            } catch (NumberFormatException nfe) {
                System.err.println("Неверный формат");
                Thread.sleep(3000);
            }
*/
            for (int i = 1; i < 10; i++) {
                progressBar.setValue(i);
                try {
                    addressLine="";
                    weight = 0.0;
                    volume = 0.0;
                    insurance = 0.0;
                    insuranceResponse = 0.0;
                    intercity = 0.0;
                    kladrFrom = "";
                    kladrTo = "";
                    longitude="";
                    latitude="";
                    coords="";
                    insuranceResponseVOZ="";
                    intercityVOZ="";
                    summaVOZ="";
                    priceFromVOZ="";
                    priceTOVOZ = "";

                    System.out.println(i);
                    Cell cell = sheet.getCell(1, i);
                    from = cell.getContents();

                    System.out.println(from);
                    System.out.print(getKladrAndSetCoords(from));
                    kladrFrom = getKladrAndSetCoords(from) + "000000000000";

                    //   setCoords(from);
                    String lat1=latitude;
                    String long1=longitude;


                    cell = sheet.getCell(2, i);
                    to = cell.getContents();
                    kladrTo = getKladrAndSetCoords(to) + "000000000000";

                    //    setCoords(to);
                    String lat2=latitude;
                    String long2=longitude;

                    //     System.out.print(getKladr(to));
                    //  Thread.sleep(10000);
                    //     System.out.println(response);


                    //  System.out.print(toCode);

                    // System.out.println("tocode= " + toCode);

                    //// cell = sheet.getCell(27, row);
                    //  cost = Integer.parseInt(cell.getContents());

                    //Thread.sleep(1000);

                    cell = sheet.getCell(10, i); //ves
                    //    System.out.println("is= " + cell.getContents().toString().replaceAll(",", "."));
                    weight = Double.parseDouble(cell.getContents().replaceAll(",", "."));

                    cell = sheet.getCell(11, i); //volume
                    //   System.out.print(cell.getContents().toString().replaceAll(",", "."));
                    volume = Double.parseDouble(cell.getContents().replaceAll(",", "."));

                    cell = sheet.getCell(32, i); //insurance
                    //   System.out.print(cell.getContents().toString().replaceAll(",", "."));
                    insurance = Double.parseDouble(cell.getContents().replaceAll(",", "."));


                    HttpPost request = new HttpPost("https://api.dellin.ru/v1/public/calculator.json");
                    StringEntity params = new StringEntity("{\"appKey\":\"8E6F26C2-043D-11E5-8F8A-00505683A6D3\",    \"derivalPoint\":\"" + kladrFrom + "\",\"derivalDoor\":true,\"arrivalPoint\":\"" + kladrTo + "\"," +
                            "\"arrivalDoor\":true,\"sizedVolume\":\"" + volume + "\",\"sizedWeight\":\"" + weight + "\",\"statedValue\":\"" + insurance + "\"}", "UTF-8");

                  /*  HttpParams httpParameters = new BasicHttpParams();
                    int timeoutConnection = 4000;
                    HttpConnectionParams.setConnectionTimeout(httpParameters, timeoutConnection);*/

            /*    String inputLine ;
                BufferedReader br = new BufferedReader(new InputStreamReader(params.getContent()));
                try {
                    while ((inputLine = br.readLine()) != null) {
                        System.out.println(inputLine);
                    }
                    br.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }

*/
                    request.addHeader("content-type", "application/json; charset=utf-8");
                    //request.addHeader("Accept-Language","ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");
                    request.setEntity(params);

                    HttpResponse response = httpClient.execute(request);
                    //   System.out.println(response);
                    //response.setHeader("Accept-Language", "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4");


                    HttpEntity entity = response.getEntity();
                   // EntityUtils.toString(sb, "UTF-8");
                    InputStream instream = entity.getContent();

                    BufferedReader reader = new BufferedReader(new InputStreamReader(instream, "UTF-8"));
                    StringBuilder sb = new StringBuilder();


                    String line = null;
                    try {
                        while ((line = reader.readLine()) != null) {
                            sb.append(line + "\n");
                        }
                    } catch (IOException e) {
                        e.printStackTrace();
                    } finally {
                        try {
                            instream.close();
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                    String ss = sb.toString();

                     System.out.println("RESPONSE: " + sb);
                    instream.close();

                    // Thread.sleep(90000);

                    JsonObject mainObject = parser.parse(ss).getAsJsonObject();
                    summa = mainObject.getAsJsonPrimitive("price").getAsDouble();

                    JsonObject mainObject2 = parser.parse(ss).getAsJsonObject().getAsJsonObject("derival");
                    priceFrom = mainObject2.getAsJsonPrimitive("price").getAsDouble();
                    terminalFrom=mainObject2.getAsJsonPrimitive("terminal").getAsString();

                    try {
                        JsonObject mainObject8 = parser.parse(ss).getAsJsonObject().getAsJsonObject("intercity");
                        intercity = mainObject8.getAsJsonPrimitive("price").getAsDouble();
                    } catch (Exception e) {
                        intercity = 0.0;
                    }

                    try {
                        JsonObject mainObject9 = parser.parse(ss).getAsJsonObject();
                        insuranceResponse = mainObject9.getAsJsonPrimitive("insurance").getAsDouble();
                    } catch (Exception e) {
                        insuranceResponse = 0.0;
                    }

                    JsonObject mainObject3 = parser.parse(ss).getAsJsonObject().getAsJsonObject("arrival");
                    terminalTO=mainObject2.getAsJsonPrimitive("terminal").getAsString();

                    priceTO = mainObject3.get("price").getAsDouble();

                    jxl.write.Label label0 = new jxl.write.Label(0, i, from);
                    jxl.write.Label label1 = new jxl.write.Label(1, i, to);
                    jxl.write.Number label2 = new jxl.write.Number(2, i, weight);
                    jxl.write.Number label3 = new jxl.write.Number(3, i, volume);
                    jxl.write.Number label5 = new jxl.write.Number(5, i, priceFrom);
                    jxl.write.Number label6 = new jxl.write.Number(6, i, intercity);
                    jxl.write.Number label7 = new jxl.write.Number(7, i, priceTO);
                    jxl.write.Number label8 = new jxl.write.Number(8, i, insuranceResponse);
                    jxl.write.Number label9 = new jxl.write.Number(9, i, summa);
                    jxl.write.Label label10 = new jxl.write.Label(10, i, terminalFrom);
                    jxl.write.Label label11 = new jxl.write.Label(11, i, terminalTO);
                    writableSheet.addCell(label0);
                    writableSheet.addCell(label1);
                    writableSheet.addCell(label2);
                    writableSheet.addCell(label3);
                    writableSheet.addCell(label5);
                    writableSheet.addCell(label6);
                    writableSheet.addCell(label7);
                    writableSheet.addCell(label8);
                    writableSheet.addCell(label9);
                    writableSheet.addCell(label10);
                    writableSheet.addCell(label11);



                    httpClient = HttpClientBuilder.create().build();

                    HttpPost requestVoz = new HttpPost("http://vozovoz.ru/api/v1/orders/price");
                    StringEntity paramsVoz = new StringEntity("{\"from\":{\"geo\":{\"latitude\":"+lat1+",\"longitude\":"+long1+"},\"address\":{\"value\"" +
                            ":\"г. Санкт-Петербург, Невский пр., д 1/4\",\"cityId\":\"61cb4131-1324-11e4-826b-d850e6bbb0fc\"},\"useppv\":false,\"date\":\"2015-07-07T00:00:00.000Z\",\"startTime" +
                            "\":\"1970-01-01T11:00:00.000Z\",\"endTime\":\"1970-01-01T14:00:00.000Z\",\"floor\":0,\"isFloor\":false,\"work\":false,\"lift\":false,\"terminal\":{\"i" +
                            "d\":\"d01da881-f94a-11e4-80c7-00155d903d03\"}},\"to\":{\"geo\":{\"latitude\":"+lat2+",\"longitude\":"+long2+"},\"address\":{\"value\":\"г. Москва, " +
                            "Тверская ул., д 1\",\"cityId\":\"544b4290-11ad-11e4-826a-d850e6bbb0fc\"},\"useppv\":false,\"date\":\"2015-07-08T00:00:00.000Z\",\"startTime\":\"1970-01-01T1" +
                            "4:00:00.000Z\",\"endTime\":\"1970-01-01T17:00:00.000Z\",\"floor\":0,\"isFloor\":false,\"work\":false,\"lift\":false,\"terminal\":{\"id\":\"7d6e3103-cd56-11e4-80c0-00155" +
                            "d903d03\"}},\"cargo\":{\"units\":[{\"length\":0.7,\"width\":0.4,\"height\":0.4,\"weight\":0.9,\"amount\":1,\"volume\":0.11199999999999999}],\"packages\":{\"visible\":false,\"" +
                            "bag1\":0,\"bag2\":0,\"sealPackage\":false,\"safePackage\":false,\"box1\":0,\"box2\":0,\"box3\":0,\"box4\":0,\"hardPackage\":false,\"extraPackage\":false,\"bubbleFilm\":false},\"" +
                            "correspondence\":false,\"insurance\":false,\"insuranceCost\":0,\"total\":{\"all\":{\"length\":0.7,\"height\":0.4,\"width\":0.4,\"volume\":"+volume+",\"weight\":"+weight+",\"amount\":1,\"dens" +
                            "ity\":8.04},\"gab\":{\"length\":0.7,\"height\":0.4,\"width\":0.4,\"volume\":0.11,\"weight\":0.9,\"amount\":1,\"density\":8.04},\"noGab\":{\"length\":0,\"height\":0,\"width\":0,\"v" +
                            "olume\":0,\"weight\":0,\"amount\":0,\"density\":null},\"max\":{\"length\":0.7,\"height\":0.4,\"width\":0.4,\"weight\":0.9}}},\"user\":{\"id\":\"54d9e7be227081aaeb610fec\",\"phoneNum" +
                            "ber\":\"new\",\"phoneApprovedHash\":null,\"type\":\"shipper\",\"shipper\":{\"type\":\"individual\",\"sendCode\":false,\"individual\":{\"fullname\":\"\",\"phoneNumber\":\"\",\"email\":" +
                            "\"\"},\"corporate\":{\"name\":\"\",\"inn\":\"\",\"kpp\":\"\",\"legalAddress\":\"\",\"contactFullname\":\"\",\"phoneNumber\":\"\",\"email\":\"\"}},\"consignee\":{\"type\":\"individual\"" +
                            ",\"sendCode\":true,\"individual\":{\"fullname\":\"\",\"phoneNumber\":\"\",\"email\":\"\"},\"corporate\":{\"name\":\"\",\"inn\":\"\",\"kpp\":\"\",\"legalAddress\":\"\",\"contactFullname" +
                            "\":\"\",\"phoneNumber\":\"\",\"email\":\"\"}},\"payer\":{\"type\":\"individual\",\"sendCode\":true,\"individual\":{\"fullname\":\"\",\"phoneNumber\":\"\",\"email\":\"\"},\"corporate\":{" +
                            "\"name\":\"\",\"inn\":\"\",\"kpp\":\"\",\"legalAddress\":\"\",\"contactFullname\":\"\",\"phoneNumber\":\"\",\"email\":\"\"}},\"uid\":\"f0ba528b-b116-11e4-80be-e15dd7ce905e\"},\"save\":t" +
                            "rue}","UTF-8");

//                String inputLine ;
//                BufferedReader br = new BufferedReader(new InputStreamReader(paramsVoz.getContent()));
//                try {
//                    while ((inputLine = br.readLine()) != null) {
//                        System.out.println(inputLine);
//                    }
//                    br.close();
//                } catch (IOException e) {
//                    e.printStackTrace();
//                }


                    requestVoz.addHeader("content-type", "application/json;charset=UTF-8");
                    requestVoz.setEntity(paramsVoz);

                    HttpResponse responseVoz = httpClient.execute(requestVoz);
                    //   System.out.println(responseVoz);

                    HttpEntity entityVoz = responseVoz.getEntity();
                    InputStream instreamVoz = entityVoz.getContent();

                    BufferedReader readerVoz = new BufferedReader(new InputStreamReader(instreamVoz));
                    StringBuilder sbVoz = new StringBuilder();

                    String lineVoz = "";
                    try {
                        while ((lineVoz = readerVoz.readLine()) != null) {
                            sbVoz.append(lineVoz + "\n");
                        }
                    } catch (IOException e) {
                        e.printStackTrace();
                    } finally {
                        try {
                            instreamVoz.close();
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                    // now you have the string representation of the HTML requestVoz
                 //   System.out.println("RESPONSE Vozovoz: " + sbVoz);
                    instreamVoz.close();

                    JsonParser parserr = new JsonParser();//response.toString()
                    JsonObject vozObj = parserr.parse(sbVoz.toString()).getAsJsonObject().getAsJsonObject("data");
                    try {
                        summaVOZ = vozObj.get("cost").toString();
                        summaVOZAction = vozObj.get("actionCost").toString();
                    } catch (Exception e){throw new Exception(e);}
                    JsonArray pItem = vozObj.getAsJsonArray("price");

                    try {
                        for (JsonElement user : pItem) {
                            JsonObject userObject = user.getAsJsonObject();
                            // if (userObject.get("ID").getAsString().equals("06"))
                            switch (userObject.get("ID").getAsString())
                            {
                                case "01": priceFromVOZ=String.valueOf(userObject.get("Cost")); break;
                                case "04": priceTOVOZ=String.valueOf(userObject.get("Cost")); break;
                                case "06": intercityVOZ=String.valueOf(userObject.get("Cost")); break;
                                case "10": insuranceResponseVOZ=String.valueOf(userObject.get("Cost")); break;


                                //return;
                            }
                            System.out.println(i);
                        }
                    } catch (Exception e) {
                    }




                    //Запись рез-тов в таблицу ДЛ+VOZ
                    try {

                        label13 = new jxl.write.Label(13, i, priceFromVOZ);
                        label14 = new jxl.write.Label(14, i, intercityVOZ);
                        label15 = new jxl.write.Label(15, i, priceTOVOZ);
                        label16 = new jxl.write.Label(16, i, insuranceResponseVOZ);
                        label17 = new jxl.write.Label(17, i, summaVOZ);
                        label18 = new jxl.write.Label(18, i, summaVOZAction);



                        writableSheet.addCell(label13);
                        writableSheet.addCell(label14);
                        writableSheet.addCell(label15);
                        writableSheet.addCell(label16);
                        writableSheet.addCell(label17);
                        writableSheet.addCell(label18);

                        if (count == 10) {
                            System.out.println(i);
                            count = 0;
                        } else count++;

                        //return;
                    } catch (Exception e) {
                        System.out.print("exc");
                    }


                } catch (Exception e) {
                    System.out.print("DoesntRecognized");
                    e.getMessage();
                    e.printStackTrace();
                    e.getStackTrace();

                }
            }

        } catch (Exception e) {
            System.out.print("exc2");
        } finally {
            writableWorkbook.write();
            writableWorkbook.close();
        }


    }

    public MonitoringDL() {
        super(new BorderLayout());

        //Create the log first, because the action listeners
        //need to refer to it.
        input = new JTextArea(1,7);
        progressBar = new JProgressBar();
        progressBar.setStringPainted(true);
        progressBar.setMinimum(0);




        log = new JTextArea(5,20);
        log.setMargin(new Insets(15,15,15,15));
        log.setEditable(false);
        JScrollPane logScrollPane = new JScrollPane(log);

        //Create a file chooser
        fc = new JFileChooser();

        //Create the open button.  We use the image from the JLF
        //Graphics Repository (but we extracted it from the jar).
        openButton = new JButton("Open a File...");
 //       openButton.addActionListener(this);
        //progressBar.addChangeListener();
        //Create the save button.  We use the image from the JLF
        //Graphics Repository (but we extracted it from the jar).
        //saveButton = new JButton("Save a File...");
        //saveButton.addActionListener(this);

        //For layout purposes, put the buttons in a separate panel
        JPanel buttonPanel = new JPanel(); //use FlowLayout
        buttonPanel.add(openButton);
        buttonPanel.add(input, BorderLayout.NORTH);
        buttonPanel.add(progressBar, BorderLayout.SOUTH);
        //buttonPanel.add(saveButton);

        //Add the buttons and the log to this panel.
        add(buttonPanel, BorderLayout.PAGE_START);
        add(logScrollPane, BorderLayout.CENTER);
        progressBar.setMaximum(enteredNumber);
        log.append("Введите количество строк в поле ввода, а затем выберите исходный файл");
    }

   /* public void actionPerformed(ActionEvent e) {

        log.append("Пожалуйста, подождите некоторое время..");
        int returnVal = fc.showOpenDialog(MonitoringDL.this);

        if (returnVal == JFileChooser.APPROVE_OPTION) {
            File file = fc.getSelectedFile();
            path= file.getAbsolutePath();
            enteredNumber=Integer.parseInt(input.getText());

            try {
                sendGet();
            } catch (Exception e1) {
                e1.printStackTrace();
            }


            log.append("Готово. Откройте файл output.xls  " + file.getName() + "." + newline);
        } else {

        }
        log.setCaretPosition(log.getDocument().getLength());

    }*/

    private static void createAndShowGUI() {

        //Create and set up the window.
        JFrame frame = new JFrame("FileChooserDemo");
        frame.setDefaultCloseOperation(EXIT_ON_CLOSE);

        //Add content to the window.
        frame.add(new MonitoringDL());
        Dimension dimension = Toolkit.getDefaultToolkit().getScreenSize();
        int x = (int) ((dimension.getWidth() - frame.getWidth()) / 2);
        int y = (int) ((dimension.getHeight() - frame.getHeight()) / 2);
        frame.setLocation(x, y);
        //Display the window.
        frame.pack();
        frame.setVisible(true);
    }




    public static void main(String[] args) throws Exception {
        MonitoringDL s= new MonitoringDL();
        s.sendGet();

      /*  SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                //Turn off metal's use of bold fonts
                UIManager.put("swing.boldMetal", Boolean.FALSE);

              //  createAndShowGUI();

            }

        }
        );*/
    }
}