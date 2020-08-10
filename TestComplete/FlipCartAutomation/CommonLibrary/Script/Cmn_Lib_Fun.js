
//Author: Santhosh
//To launch the chrome browser and Run the URL (Still I can write in descriptive way if i get moretime)
function LaunchBrowser(BrowserType, pageURL)
{
  
  //Laumch browsers
  Browsers.Item(BrowserType).Run(pageURL);
  
}
  
 
//Author: Santhosh
//I will write the descriptive function like pageobject(refer last function in the page) for each objects in below function
 function logingFlipKart(username,PW)
 {
   
   Sys.Browser("chrome").Page("https://www.flipkart.com/").Panel(0).Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(0).Form(0).Panel(0).Textbox(0).SetText(username);
   Sys.Browser("chrome").Page("https://www.flipkart.com/").Panel(0).Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(0).Form(0).Panel(1).PasswordBox(0).SetText(PW);
   Sys.Browser("chrome").Page("https://www.flipkart.com/").Panel(0).Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(0).Form(0).Panel(2).Button(0).ClickButton();
 }
 
//Author: Santhosh
//I will write the descriptive function like pageobject(refer last function in the page) for each objects in below function
 function searchItem(sItem)
  {
  Delay(3000);  
  Sys.Browser("chrome").Page("https://www.flipkart.com/").Panel("container").Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Form(0).Panel(0).Panel(0).Textbox("q").Keys(sItem);
  Sys.Browser("chrome").Page("https://www.flipkart.com/").Panel("container").Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Form(0).Panel(0).Button(0).ClickButton();
  Sys.Browser("chrome").Page("https://www.flipkart.com/search*").Panel("container").Panel(0).Panel(2).Panel(1).Panel(0).Panel(1).Panel(1).Panel(0).Panel(0).Panel(0).Link(0).Click();
  }
  
//Author: Santhosh
//I will write the descriptive function like pageobject(refer last function in the page) for each objects in below function
function AddToCart()
{
  Delay(3000);  
  Sys.Browser("chrome").Page("https://www.flipkart.com/*").Panel("container").Panel(0).Panel(2).Panel(0).Panel(0).Panel(1).Panel(0).Button(0).ClickButton()

}


//Author: Santhosh
//Common function to read the data from excel file and assign them to projectsuite level variables as global

//Similar way we can write the javascripts to read the *****json***** file data and assign them into variables
function ReadInputFromExcel()
{
  var excel,excelbook,excelsheet,rowno,columnno;
  
  excel = new ActiveXObject("Excel.Application");
  excelBook = excel.Workbooks.Open("C:\\Users\\PAVAN\\Desktop\\Santhosh\\JavaScripts\\TestExcel.xlsx");
  excelsheet = excel.Worksheets("Sheet1");
  rowNo = excelsheet.UsedRange.Rows.Count;
  ColNo = excelsheet.UsedRange.Columns.Count;
  //First row of the sheet is columnname sheet and column names are fixed
  for(i = 1; i <= rowNo;i++) 
  {
    ProjectSuite.Variables.BrowserType = excelsheet.Cells(1, i).value;
    ProjectSuite.Variables.URL = excelsheet.Cells(2, i).value;
    ProjectSuite.Variables.FlipkartUN = excelsheet.Cells(3, i).value;
    ProjectSuite.Variables.FlipkartPW = excelsheet.Cells(4, i).value;
    ProjectSuite.Variables.FlipkartItem = excelsheet.Cells(5, i).value;
  }

  excelBook.Save();
  excel.Quit();
}




//Author: Santhosh
//function to identify the page by passing browser type and page url
//Below is the common library function using javascripts(Can be use it for any page)
//Similar way I can write common functions for editbox, button, dropdowns, Links, etc - 
//(Just by passing page name, object type and unique object property value we cab retrun the object or null if not found)
function pageobject(BrowserType, pageURL)
{
  var brow_obj, pageObj;
  //Assign the value from global varibles list
  BrowserType = ProjectSuite.Variables.BrowserType;
  pageURL = ProjectSuite.Variables.URL;
  
  brow_obj = Utils.CreateStubObject();
  brow_obj = Sys.Browser(BrowserType);
  if (brow_obj.Exists)
  {
    pageObj = brow_obj.FindChild("url",URL +"*");  //regular expression is used
    
    if (pageObj.Exist){return pageObj;}
      else {return null;}
  }
  
}




