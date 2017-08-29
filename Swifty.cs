using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Drawing;
using System.Text;
using System.Threading;
using System.IO;
using System.Windows.Forms;

namespace Swifty
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {

        }

        public static IWebDriver driver;
        public static string smartITUser;
        public static string smartITPassword;
        public bool submitFlag = false;
        public static string assetAssignee;
        public static string assetNumber;
        public static bool end;

        public void Execute_Step(string step, int count) // Executes steps for all workflow routines
        {
            try
            {
                Log_Text_Box.AppendText(DateTime.Now.ToString("[M/d/yyyy @ h:mm:stt] ") + "Executing step: "+ step + "\r\n");

                if (!step.Substring(0,2).Equals("//"))
                {
                    SendKeys.Send(step);
                    SendKeys.Send("{ENTER}");
                }
                else
                {
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                    wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath($"{step}")));

                    IWebElement clickEvent = driver.FindElement(By.XPath($"{step}"));

                    clickEvent.Click();
                }
            }
            catch (WebDriverException) // Retries up to three times before quitting
            {
                count++;
                if (count == 4)
                {
                    Log_Text_Box.AppendText(DateTime.Now.ToString("[M/d/yyyy @ h:mm:stt] ") + $"Step {step} has failed." + "\r\n");
                    driver.Quit();
                    return;
                }
                else
                {
                    Log_Text_Box.AppendText(DateTime.Now.ToString("[M/d/yyyy @ h:mm:stt] ") + $"Could not execute step: {step}." + "\r\n");
                    Log_Text_Box.AppendText(DateTime.Now.ToString("[M/d/yyyy @ h:mm:stt] ") + $"Attempting step: {step} again." + "\r\n");
                    Execute_Step(step, count);
                }
            }
        }

        public void Asset_Update() // Executes asset update steps. (Generic steps used as placholder for redacted content)
        {
            string[] steps =
            {
                "//a[contains(text(), 'Console')]",
                "//span[contains(text(), 'Asset Console')]",
                "//button[contains(text(), 'Clear Filters')]",
                "//span[contains(text(), 'Filter')]",
                "//div[contains(text(), 'Keywords')]",
                "//input[contains(@placeholder, 'Type a specific search term')]",
                assetNumber,
                "//div[contains(text(), 'Asset Type')]",
                "//div[contains(text(), 'Computer System')]",
                "//div[contains(@ng-style, 'rowStyle(row)')]",
                "//div[contains(text(), 'Status:')]",
                "//div[contains(@title-text, 'Asset Status')]",
                "//a[contains(text(), 'Deployed')]",
                "//button[contains(text(), 'Save')]",
                "//a[contains(text(), 'People')]",
                "//div[contains(@ng-click, 'addRelatedPeople()')]",
                "//input[contains(@placeholder, 'Type to search people')]",
                assetAssignee,
                "//div[contains(@ng-click, 'selectPerson(person)')]",
                "//button[contains(text(), 'Add People')]"
            };

            foreach (string i in steps)
            {
                Execute_Step(i, 0);
            }

            // Takes confirmation screenshot.
            Thread.Sleep(3000);
            driver.Manage().Window.Size = new Size(1920, 1080);
            SaveScreenshot(driver, assetAssignee);

            AU_Process_Button.Enabled = true;
            driver.Quit();
        }

        public void Repairs() // Executes repair steps. (Generic steps used as placholder for redacted content)
        {
            string[] steps =
            {
                "//a[contains(text(), 'Console')]",
                "//span[contains(text(), 'Asset Console')]",
                "//button[contains(text(), 'Clear Filters')]",
                "//span[contains(text(), 'Filter')]",
                "//div[contains(text(), 'Keywords')]",
                "//input[contains(@placeholder, 'Type a specific search term')]",
                assetNumber,
                "//div[contains(text(), 'Asset Type')]",
                "//div[contains(text(), 'Computer System')]",
                "//div[contains(@ng-style, 'rowStyle(row)')]",
                "//div[contains(text(), 'Status:')]",
                "//div[contains(@title-text, 'Asset Status')]",
                "//a[contains(text(), 'Deployed')]",
                "//button[contains(text(), 'Save')]",
                "//a[contains(text(), 'People')]",
                "//div[contains(@ng-click, 'addRelatedPeople()')]",
                "//input[contains(@placeholder, 'Type to search people')]",
                "//div[contains(@ng-click, 'selectPerson(person)')]",
                "//button[contains(text(), 'Add People')]"
            };

            foreach (string i in steps)
            {
                Execute_Step(i, 0);
            }

            // Takes confirmation screenshot.
            Thread.Sleep(3000);
            driver.Manage().Window.Size = new Size(1920, 1080);
            SaveScreenshot(driver, assetNumber);

            R_Process_Button.Enabled = true;
            driver.Quit();
        }

        public void Credential_Input_Check() // Validates ID format before processing requests.
        {
            if (submitFlag == true)
            {
                ID_Label.Visible = true;
                Password_Label.Visible = true;
                ID_Textbox.Visible = true;
                Password_Textbox.Visible = true;

                Submit_Credentials_Button.Text = "Submit";
                submitFlag = false;
            }
            else
            {
                if (string.IsNullOrEmpty(ID_Textbox.Text) || string.IsNullOrEmpty(Password_Textbox.Text))
                {
                    MessageBox.Show("Enter your network ID and password.");
                }
                else if (ID_Textbox.Text.Length != 7)
                {
                    MessageBox.Show("Corporate ID is not correctly formatted.");
                }
                else
                {
                    smartITUser = ID_Textbox.Text;
                    smartITPassword = Password_Textbox.Text;

                    ID_Label.Visible = false;
                    Password_Label.Visible = false;
                    ID_Textbox.Visible = false;
                    Password_Textbox.Visible = false;

                    Submit_Credentials_Button.Text = smartITUser;
                    submitFlag = true;
                }
            }
        }

        public static void SaveScreenshot(IWebDriver driver, string ScreenShotFileName) // Takes screenshot of completed task for visual confirmation.
        {
            var folderLocation = ("C:\\Users\\%USERPROFILE%\\Desktop\\Screenshots\\");

            if (!Directory.Exists(folderLocation))
            {
                Directory.CreateDirectory(folderLocation);
            }

            var screenShot = ((ITakesScreenshot)driver).GetScreenshot();
            var fileName = new StringBuilder(folderLocation);

            fileName.Append(ScreenShotFileName);
            fileName.Append(DateTime.Now.ToString("_M-d-yyyy_h:mm:stt"));
            fileName.Append(".png");
            screenShot.SaveAsFile(fileName.ToString(), ScreenshotImageFormat.Png);
        }

        public void Driver_Setup() // Sets up web driver and logs user in
        {
            string webURL = "https://example.com";

            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;

            ChromeOptions options = new ChromeOptions();
            options.AddArgument("disable-infobars");
            options.AddArgument("incognito");
            options.AddArgument("start-maximized");

            driver = new ChromeDriver(service, options);
            
            driver.Navigate().GoToUrl(webURL);

            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath("//*[@id='loginUserName']")));

                IWebElement login = driver.FindElement(By.XPath("//*[@id='loginUserName']"));
                IWebElement password = driver.FindElement(By.XPath("//*[@id='loginPass']"));
                IWebElement loginButton = driver.FindElement(By.XPath("//button[contains(text(), 'Log In')]"));

                login.SendKeys(smartITUser);
                password.SendKeys(smartITPassword);
                loginButton.Click();
            }
            catch (WebDriverException)
            {
                Log_Text_Box.AppendText(DateTime.Now.ToString("[M/d/yyyy @ h:mm:stt] ") +  "Could not find login fields." + "\r\n");
                driver.Quit();
            }

        }

        private void Submit_Credentials_Button_Click(object sender, EventArgs e)
        {
            Credential_Input_Check();
        }

        private void Close_Browser_Button_Click(object sender, EventArgs e)
        {
            AU_Process_Button.Enabled = true;
            try
            {
                driver.Quit();
            }
            catch (Exception)
            {
                Log_Text_Box.AppendText(DateTime.Now.ToString("[M/d/yyyy @ h:mm:stt] ") + "There are no open browsers." + "\r\n");
            }
            
        }

        private void Clear_Log_Button_Click(object sender, EventArgs e)
        {
            Log_Text_Box.Clear();
        }

        private void Asset_Update_Process_Button_Click(object sender, EventArgs e)
        {
            if (submitFlag == true)
            {
                AU_Process_Button.Enabled = false;
                assetAssignee = Asset_Update_Assignee_Text_Box.Text;
                assetNumber = Asset_Update_Asset_Number_Text_Box.Text;
                Driver_Setup();
                Asset_Update();
            }
            else
            {
                MessageBox.Show("Please enter your network credentials.");
            }
        }

        private void Repair_Process_Button_Click(object sender, EventArgs e)
        {
            if (submitFlag == true)
            {
                R_Process_Button.Enabled = false;
                assetNumber = Repairs_Asset_Number_Text_Box.Text;
                Driver_Setup();
                Repairs();
            }
            else
            {
                MessageBox.Show("Please enter your network credentials.");
            }
        }
    }
}
