package loginjiaowuwang;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URI;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import javax.lang.model.element.Element;
import javax.swing.text.Document;

import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.NameValuePair;
import org.apache.http.client.ClientProtocolException;
import org.apache.http.client.HttpClient;
import org.apache.http.client.ResponseHandler;
import org.apache.http.client.entity.UrlEncodedFormEntity;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.mime.Header;
import org.apache.http.impl.client.BasicResponseHandler;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.http.message.BasicNameValuePair;
import org.apache.http.util.EntityUtils;
import org.jsoup.Jsoup;
import org.jsoup.select.Elements;
import org.omg.CORBA.portable.InputStream;

import jxl.Workbook;
import jxl.write.WritableWorkbook;

class clienteduSystem
{
	public static HttpClient client=new DefaultHttpClient();
	public static String TextBox1="*******";//你的学号
	public static String gettimetable="http://jwby.hyit.edu.cn/xskbcx.aspx?xh="+TextBox1;
	public static String getscoreurl="http://jwby.hyit.edu.cn/xscj_gc.aspx?xh="+TextBox1;
	public static String button="按成绩查询";
	public static List<String>  getJESSID() throws IOException
	{
		String url="http://jwby.hyit.edu.cn/";
		org.jsoup.nodes.Document doc=Jsoup.connect(url).get();
		Elements element=doc.select("form").select("input[name=__VIEWSTATE]");
		String view=(String)element.attr("value");
		String TextBox2="******";//你的密码
		String RadioButtonList="学生";
		String TextBox3="";
		String Button1="";
		List<String> list1=new ArrayList<String>();
		List<NameValuePair> list=new ArrayList<NameValuePair>();
		list.add(new BasicNameValuePair("__VIEWSTATE",view));
		list.add(new BasicNameValuePair("TextBox1",	TextBox1));
		list.add(new BasicNameValuePair("TextBox2",TextBox2));
		list.add(new BasicNameValuePair("TextBox3",TextBox3));
		list.add(new BasicNameValuePair("RadioButtonList1",RadioButtonList));
		list.add(new BasicNameValuePair("Button1",Button1));
		HttpPost post=new HttpPost(url);
		try
		{
			post.setEntity(new UrlEncodedFormEntity(list));
			HttpResponse response=client.execute(post);
		//	System.out.println("response"+response);
		//	System.out.println("status:"+response.getStatusLine());
			HttpEntity entity=response.getEntity();
			//System.out.println("entity"+entity);
			String result=EntityUtils.toString(entity,"utf-8");
			//System.out.println("result"+result);
			org.apache.http.Header location=response.getFirstHeader("Location");
			String s=location.getValue().toString();
			list1.add(s);
			org.apache.http.Header cookie=response.getFirstHeader("Set-Cookie");
			String cookie1=cookie.getValue().toString();
			list1.add(cookie1);
			return list1;
		}catch(Exception e)
		{
			e.printStackTrace();
			return null;
		}
		finally
		{
			post.abort();
		}
		
		
	}
	public static void createxcel(String score[][])
	{
		String [][] s=new String[100][2];
		int num=0;
		int j=0;
		for(int i=0;i<100;i++)
		{
		
				if(score[i][0]!=null)
				{	
					s[j][0]=score[i][0];
					s[j++][1]=score[i][1];
					num++;
				}
				
				
		}

		HSSFWorkbook workbook=new HSSFWorkbook();
		HSSFSheet sheet=workbook.createSheet("成绩表");
		HSSFRow row=sheet.createRow(0);
		HSSFCell cell=row.createCell(0);
		cell.setCellValue(s[num-1][0]);
		cell=row.createCell(2);
		cell.setCellValue(s[num-1][1]);
		for(int k=0;k<num-1;k++)
		{
			HSSFRow row1=sheet.createRow(k+1);
			row1.createCell(0).setCellValue(s[k][0]);
			row1.createCell(2).setCellValue(s[k][1]);
			
		}
		try
		{
			FileOutputStream file=new FileOutputStream("E:\\文件\\教务网\\成绩表.xls");
			workbook.write(file);
			System.out.println("写入成功");
			file.close();
		}catch(IOException e)
		{
			e.printStackTrace();
		}
	}
	public static void gethtml(String url1,String cookie) throws IOException
	{

		HttpGet get=new HttpGet(url1);
		
		get.setHeader("Referer",url1);  
		ResponseHandler<String> responsehandler=new BasicResponseHandler();
		String responsebody;

	try
	{
		responsebody=client.execute(get,responsehandler);
		//System.out.println(responsebody);
		//System.out.println(get.getFirstHeader("Host"));
	}catch (Exception e) {  
        e.printStackTrace();  
        responsebody = null;  
    } finally {  
        get.abort();   
    }  
}
	public static String getview(String url1) throws IOException
	{
		HttpGet get=new HttpGet(url1);
		get.setHeader("Referer",url1);
		HttpResponse response=client.execute(get);
		org.jsoup.nodes.Document doc=Jsoup.parse(EntityUtils.toString(response.getEntity()));
		String view=doc.select("form").select("input[name=__VIEWSTATE]").attr("value").toString();
		return view;
    }  
	public static String getscore(String url) throws IOException
	{
		String view=getview(url);
		System.out.println(view);
		HttpPost post=new HttpPost(url);
		post.setHeader("Referer", url);
		ResponseHandler<String> responsehandler=new BasicResponseHandler();
		String button="按学期查询";
		String responsebody="";
		List<NameValuePair>list=new ArrayList<NameValuePair>();
		list.add(new BasicNameValuePair("__VIEWSTATE",view));
		list.add(new BasicNameValuePair("ddlXN",""));
		list.add(new BasicNameValuePair("ddlXQ",""));
		list.add(new BasicNameValuePair("Button1",button));
		post.setEntity(new UrlEncodedFormEntity(list));
		responsebody=client.execute(post,responsehandler);
		return responsebody;
	}
	public static String [][] getscoreinformation(String html)
	{
		int i=3;
		int j=0;
		int k=0;
		int l=0;
		int t=0;
		int b=0;
		String[][] main=new String[1000][1000];
		String[][] main3=new String[1000][1000];
		String [] main1=new String[10000];
		String [] main2=new String[10000];
		String[] name=new String[10];
		org.jsoup.nodes.Document doc=Jsoup.parse(html);
		Elements element=doc.select("body").select("form").select("div.toolbox").select("div.searchbox").select("p.search_con").select("span");
			for(i=3;i<9;i++)
				name[j++]=element.select("#Label"+i).text();
			for(i=0;i<j;i++)
System.out.println(name[i]);
		
			Elements elementmain=doc.select("body").select("form").select("div[class=main_box]").select("div").select("span").select("fieldset").select("table[id=Datagrid1]").select("tbody").select("tr");
			Elements elementmainart=doc.select("body").select("form").select("div[class=main_box]").select("div").select("span").select("fieldset").select("table[id=TabTj]").select("tbody").select("tr").select("td").select("table[id=DataGrid6]").select("tbody").select("tr");
			for(org.jsoup.nodes.Element element2:elementmain)
			{
			Elements element3=element2.select("td");
					for(org.jsoup.nodes.Element element4:element3)
					{
						String s=element4.text();
						main1[l++]=s;
						
					}
		
					
			}
			for(org.jsoup.nodes.Element element5:elementmainart)
			{
			Elements element6=element5.select("td");
					for(org.jsoup.nodes.Element element7:element6)
					{
						String s=element7.text();
						main2[t++]=s;
						
					}
		
					
			}
			int m=3;
			for(i=0;i<l/10;i++)
			{
				k=m;
				j=k+5;
				main[i][0]=main1[k];
				main[i][1]=main1[j];
				k=j+10;
				m=k;
			}
			for(i=0;i<t;i++)
			{
				System.out.println(main2[i]);
			}
			int s=0;
			for(i=0;i<t/5;i++)
			{
				k=s;
				j=k+2;
				main3[i][0]=main2[k];
				main3[i][1]=main2[j];
				k=j+3;
				s=k;
			}	
			for(i=l/10-1;i<l/10+t/5;i++)
			{
				main[i][0]=main3[b][0];
				main[i][1]=main3[b][1];
				b++;
			}
			for(i=l/10+t/5;i<l/10+t/5+1;i++)
			{
				main[i][0]=name[0];
				main[i][1]=name[2];
			}
			for(i=0;i<l/10+t/5+1;i++)
			{
				
				if(main[i][0]!=null)
				{
					System.out.print(main[i][0]);
					System.out.println(main[i][1]);
				}
				
			}
			return main;
	}
	public static void main(String []args) throws IOException
	{
		List<String> location=getJESSID();
		String[][] score=new String[1000][1000];
		String url="http:"+"//"+"jwby.hyit.edu.cn"+location.get(0);
		String cookie=location.get(1);
		String scorehtml=getscore(getscoreurl);
		score=getscoreinformation(scorehtml);
		createxcel(score);
	}
}
