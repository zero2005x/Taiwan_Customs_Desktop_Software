import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.*;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.WebElement;
import java.util.*;
public class Find_Main_Appraiser {
	
	public WebDriver driver = new ChromeDriver();
	public String Application_Declaration_Number = "";
	public JavascriptExecutor js = (JavascriptExecutor)driver;
	public String Max_taxPayMethodR = "";
	public int Max_Value = 0, item_number = 0;
	public Find_Main_Appraiser(String Target) {
		driver.get("http://aci.customs.gov.tw/portal/Login_main");
        driver.findElement(By.name("user.userId")).sendKeys("");
        driver.findElement(By.name("user.password")).sendKeys("" + Keys.ENTER);
        String Target_URL = "http://aci.customs.gov.tw/APIM/";
        Target_URL += Target;
        Target_URL += "?opener=true&?cust_Cd=AW";
        driver.get(Target_URL);
        
	}
	
	public void Application_Declaration_Number_Transformation() {
		int Number_Length = Application_Declaration_Number.length();
		if(Number_Length == 12) {
			String temp = Application_Declaration_Number.substring(0, 3)
			+ "  "
			+ Application_Declaration_Number.substring(3, 12);
			Application_Declaration_Number = temp;
			System.out.println(temp);
					
		}
		else if(Number_Length == 14) {
			System.out.println(Application_Declaration_Number);
		}
		else 
			Application_Declaration_Number = "";
		return;
	}
	
	public void Enter_Application_Declaration_Number() {
		System.out.println("Enter Application Declaration Number:");
		Scanner Input = new Scanner(System.in);
		Application_Declaration_Number = Input.nextLine();
		driver.findElement(By.name("declNo1")).sendKeys(Application_Declaration_Number + Keys.ENTER);
		js.executeScript("f6();");
        
		Input.close();
		return;
	}
	
	public void Pressing_Next_Page() {
		js.executeScript("forward();");
		return;
	}
	
	int total_page_number() {		
		String Total_Page_Number = "";
		while(Total_Page_Number == "") {
			Total_Page_Number = js.executeScript("return document.querySelector(\"#total\").innerText;").toString();
		}
		System.out.println(Total_Page_Number);
		return Integer.parseInt(Total_Page_Number);
		
	}
	
	public String Find_taxPayMethod(int index) {
		String taxPayMethod = "";
		String taxPayMethod_xpath = "/html/body/div[2]/div[3]/form/div[1]/div[1]/table/tbody/tr[3]/td/div[2]/table/tbody/tr[";
		taxPayMethod_xpath += Integer.toString(index*2 + 1);
		taxPayMethod_xpath +=  "]/td[2]/select";
		WebElement select = driver.findElement(By.xpath(taxPayMethod_xpath));
		Select dropDown = new Select(select);           
		String Select = dropDown.getFirstSelectedOption().getText().toString();
		
        if(Select != "") 
        	taxPayMethod = Select.substring(0, 2);
        System.out.println(taxPayMethod);
		return taxPayMethod;
		
	}
	
	public String Find_Target_Value(int index) {
		String Item_Value = "";
		String Executable_Javascript_Commend_Left = "return $(\"input[name='custValAmtR']\").eq(";
        String Executable_Javascript_Commend_Combine = "";
        Executable_Javascript_Commend_Combine += Executable_Javascript_Commend_Left;
        Executable_Javascript_Commend_Combine += Integer.toString(index) + ").val();";
 
        int count = 0;
        while(Item_Value == "") {
        	Item_Value = js.executeScript(Executable_Javascript_Commend_Combine).toString();
        	if(Item_Value == "")
        		count += 1;
        	if(count > 300)
        		break;
        }
        if(Item_Value != "") 
        	Item_Value = Item_Value.replaceAll(",", "");
        
        System.out.println(Item_Value);
		return Item_Value;
	}

	public void check_status() {
		String Status = "";
		while(Status.compareTo("查詢成功!") != 0) 
			Status = js.executeScript("return $(\"#statusMsg\").val();").toString();
		System.out.println(Status);
		return;
	}
	
	public static void main(String[] args) {
		Find_Main_Appraiser aba= new Find_Main_Appraiser("IE07");
        aba.Enter_Application_Declaration_Number();
        int total_page_number = aba.total_page_number(), current_value = 0, item_number_count = 1;
        String  Item_Value_String = "";
        Boolean Last_Item = false;
        for(int page = 0 ; page < total_page_number; page++) {
        	aba.check_status();
        	if (Last_Item == true)
        		break;
        	for(int index = 1; index < 21 ; index++){
        		Item_Value_String =aba.Find_Target_Value(index);
        		if(aba.Find_Target_Value(index) == "") {
        			Last_Item =true;
        			break;
        		}
        		int Item_Value = Integer.valueOf(Item_Value_String);
        		item_number_count += 1;
        		if(current_value < Item_Value) {
        			current_value = Item_Value;
        			aba.Max_Value = Item_Value;
        			aba.Max_taxPayMethodR = aba.Find_taxPayMethod(index);
        			aba.item_number = item_number_count;
        		}
        	}
        	if(page < total_page_number - 1) {
        		aba.Pressing_Next_Page();
        		System.out.println("Changing Page");	
        	}     		
        }
        System.out.println(aba.item_number);
        System.out.println(aba.Max_Value);
        System.out.println(aba.Max_taxPayMethodR);  
	}
}
