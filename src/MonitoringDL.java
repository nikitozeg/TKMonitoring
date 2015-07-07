import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.HttpClientBuilder;
import org.openqa.jetty.util.URI;

import java.io.*;

public class MonitoringDL {
   // File exlFile = new File("input.xls");
   File exlFile = new File("C:\\Users\\n.ivanov\\Dropbox\\AutoMonitoringDL\\input220volt2.xls");
    Workbook w;
    String fromCode, toCode, insuranceResponse, intercity, kladrFrom,kladrTo;
    Double weight,volume, insurance;
    int count;


    public static void main(String[] args) throws Exception {
        MonitoringDL http = new MonitoringDL();
        System.out.println("Testing 1 - Send Http GET request");
        http.sendGet();

    }

    private String getCoordinates(String address)throws Exception
    {
        String kladr="";
        HttpClient httpClient1 = HttpClientBuilder.create().build();
        HttpPost request1 = new HttpPost("https://geocode-maps.yandex.ru/1.x/?format=json&geocode="+URI.encodePath("address"));

        HttpResponse response1 =  httpClient1.execute(request1);
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
        kladr=mainObject.get(0).getAsJsonObject().getAsJsonObject("data").get("kladr_id").getAsString();

        return kladr;

    }


    private String getKladr(String address)throws Exception
    {
        String kladr="";
        HttpClient httpClient1 = HttpClientBuilder.create().build();
        HttpPost request1 = new HttpPost("https://suggestions.dadata.ru/suggestions/api/4_1/rs/suggest/address");
        StringEntity params1 = new StringEntity("{\"count\":1,\"query\":\""+address+"\"}","utf-8");
        request1.addHeader("content-type", "application/json");
        request1.addHeader("Authorization", "Token 84beb76a98914195f374779f2f313d31efca3c5d");
        request1.addHeader("X-Secret", "cb82deee2d367b967ba569b5fc11b9e21a8c4832");
        request1.setEntity(params1);

        HttpResponse response1 =  httpClient1.execute(request1);
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
        kladr=mainObject.get(0).getAsJsonObject().getAsJsonObject("data").get("kladr_id").getAsString();

        return kladr;

    }


    public void sendGet() throws Exception {
        String summa, derival, priceFrom, fromm, priceTO = "";
     //   File crowlerResult = new File("output.xls");
        File crowlerResult = new File("C:\\Users\\n.ivanov\\Dropbox\\AutoMonitoringDL\\output.xls");
        w = Workbook.getWorkbook(exlFile);
        Sheet sheet = w.getSheet(0);
        WritableWorkbook writableWorkbook = Workbook.createWorkbook(crowlerResult);
        WritableSheet writableSheet = writableWorkbook.createSheet("Sheet2", 0);
        Label label00 = new Label(2, 0, "МТ+Забор+Отвоз");
        Label label01 = new Label(3, 0, "Забор");
        Label label02 = new Label(4, 0, "Отвоз");
        Label label03 = new Label(4, 0, "Отвоз");
        Label label04 = new Label(6, 0, "Вес");
        writableSheet.addCell(label01);
        writableSheet.addCell(label02);
        writableSheet.addCell(label00);
        writableSheet.addCell(label03);
        writableSheet.addCell(label04);

        try {
            int enteredNumber=0;
            String to = ""; String from = "";

            BufferedReader br = new BufferedReader(new InputStreamReader(System.in));
            System.out.print("Введите количество обрабатываемых строк:");
            try{
                 enteredNumber = Integer.parseInt(br.readLine());
            }catch(NumberFormatException nfe){
                System.err.println("Неверный формат");
                Thread.sleep(3000);
            }

            for (int i = 1; i < enteredNumber; i++) {
                try {
                    weight=0.0;
                    volume=0.0;
                    insurance=0.0;
                    insuranceResponse="";
                    intercity="";
                    kladrFrom="";
                    kladrTo="";

                    System.out.println(i);
                    Cell cell = sheet.getCell(1, i);
                    from = cell.getContents();

                    System.out.println(from);
                    System.out.print(getKladr(from));
                    kladrFrom=getKladr(from)+"000000000000";


                    cell = sheet.getCell(2, i);
                    to = cell.getContents();
                    kladrTo=getKladr(to)+"000000000000";
                    System.out.print(getKladr(to));
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





                    HttpClient httpClient = HttpClientBuilder.create().build();

                    HttpPost request = new HttpPost("https://api.dellin.ru/v1/public/calculator.json");
                    StringEntity params = new StringEntity("{\"appKey\":\"8E6F26C2-043D-11E5-8F8A-00505683A6D3\",    \"derivalPoint\":\"" + kladrFrom + "\",\"derivalDoor\":true,\"arrivalPoint\":\"" + kladrTo + "\"," +
                            "\"arrivalDoor\":true,\"sizedVolume\":\""+volume + "\",\"sizedWeight\":\"" + weight + "\",\"statedValue\":\"" + insurance+ "\"}");

              /*  String inputLine ;
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

                    request.addHeader("content-type", "application/javascript");
                    request.setEntity(params);

                    HttpResponse response = httpClient.execute(request);
                    //   System.out.println(response);

                    HttpEntity entity = response.getEntity();
                    InputStream instream = entity.getContent();

                    BufferedReader reader = new BufferedReader(new InputStreamReader(instream));
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
                    // now you have the string representation of the HTML request
                    // System.out.println("RESPONSE: " + ss);
                    instream.close();

                    // Thread.sleep(90000);

                    JsonParser parser = new JsonParser();//response.toString()

                    JsonObject mainObject = parser.parse(ss).getAsJsonObject();
                    summa = mainObject.getAsJsonPrimitive("price").getAsString();
                    //   System.out.println("intercity= " + summa);

                    JsonObject mainObject2 = parser.parse(ss).getAsJsonObject().getAsJsonObject("derival");
                    priceFrom = mainObject2.getAsJsonPrimitive("price").getAsString();

                    try {
                        JsonObject mainObject8 = parser.parse(ss).getAsJsonObject().getAsJsonObject("intercity");
                        intercity = mainObject8.getAsJsonPrimitive("price").getAsString();
                    }catch (Exception e) { intercity = "-";}

                    try {
                        JsonObject mainObject9 = parser.parse(ss).getAsJsonObject();
                        insuranceResponse = mainObject9.getAsJsonPrimitive("insurance").getAsString();
                    }catch (Exception e) { insuranceResponse = "-";}
                    //   System.out.println("priceFrom= " + priceFrom);

                    fromm = mainObject2.get("terminal").getAsString();
                    //   System.out.println("from= " + fromm);

                    JsonObject mainObject3 = parser.parse(ss).getAsJsonObject().getAsJsonObject("arrival");
                    derival = mainObject3.get("terminal").getAsString();
                    //   System.out.println("to= " + derival.toString());

                    priceTO = mainObject3.get("price").getAsString();
                    //     System.out.println("priceTO= " + priceTO);

                    //     System.out.println("weight= " + weight);

                    try {
                        // for (JsonElement user : pItem) {
                        //   JsonObject userObject = user.getAsJsonObject();
                        //if (userObject.get("type").getAsString().equals("avto")) {
                        //        System.out.print(userObject.get("price"));

                        Label label0 = new Label(0, i, from);
                        Label label1 = new Label(1, i, to);
                        Label label2 = new Label(2, i, summa);
                        Label label3 = new Label(3, i, priceFrom);
                        Label label4 = new Label(4, i, priceTO);
                        Label label5 = new Label(6, i, weight.toString());
                        Label label6 = new Label(7, i, volume.toString());
                        Label label7 = new Label(8, i, insuranceResponse);
                        Label label8 = new Label(9, i, intercity);

                        writableSheet.addCell(label0);
                        writableSheet.addCell(label1);
                        writableSheet.addCell(label2);
                        writableSheet.addCell(label3);
                        writableSheet.addCell(label4);
                        writableSheet.addCell(label5);
                        writableSheet.addCell(label6);
                        writableSheet.addCell(label7);
                        writableSheet.addCell(label8);
                        if (count==10){
                            System.out.println(i);count=0;}
                        else count++;

                        //return;
                    } catch (Exception e) {System.out.print("exc");
                    }


                } catch (Exception e) {
                    Label label0 = new Label(0, i, "Моск Обл");
                    writableSheet.addCell(label0);
                    System.out.print("DoesntRecognized");
                }
            }

        } catch (Exception e) {System.out.print("exc2");
        } finally {
            writableWorkbook.write();
            writableWorkbook.close();
        }


    }

}
