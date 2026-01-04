//原神Webイベントチェッカー

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using IniParser;

namespace GenshinEventChecker
{
    class Program
    {
        static void Main(string[] args)
        {
            var ini = new FileIniDataParser().ReadFile("./setting.ini");

            var start_url = "https://www.hoyolab.com/circles/2/27/official?page_type=27&page_sort=events";
            var driver_dir = ini["Directory"]["driver"].Trim('"');
            var binary_dir = ini["Directory"]["binary"].Trim('"');
            var user_dir = ini["Directory"]["user"].Trim('"');

            var service = ChromeDriverService.CreateDefaultService(driver_dir);
            service.HideCommandPromptWindow = true;
            var options = new ChromeOptions();
            options.BinaryLocation = binary_dir;

            //前回セッション取得
            Dictionary<string, object> session;
            try
            {
                using (var reader = new StreamReader(user_dir + "/Default/Preferences", Encoding.UTF8))
                {
                    var domein_ary = start_url.Replace(".com", ".com_").Replace(".jp", ".jp_").Split(new string[] { "://", "?" }, StringSplitOptions.None)[1].Split('.');
                    var token = JObject.Parse(reader.ReadToEnd())["browser"]["app_window_placement"];
                    foreach (var domein in domein_ary)
                    {
                        token = token[domein];
                    }
                    session = new Dictionary<string, object>
                    {
                        { "left", (int)token["left"] },
                        { "right", (int)token["right"] },
                        { "top", (int)token["top"] },
                        { "bottom", (int)token["bottom"] },
                        { "maximized", (bool)token["maximized"] }
                    };
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                session = new Dictionary<string, object>
                {
                    { "left", 0 },
                    { "right", 500 },
                    { "top", 0 },
                    { "bottom", 500 },
                    { "maximized", false }
                };
            }

            //起動オプション
            var arguments = new string[] {
                "--user-data-dir=" + user_dir, 
                "--profile-directory=Default",
                "--app=" + start_url, //アプリケーションモードで起動
                //"--start-maximized", //最大化(--window-position, --window-sizeと併用不可)
                "--window-size=" + ((int)session["right"] - (int)session["left"]).ToString() + ','+((int)session["bottom"] - (int)session["top"]).ToString(), //ウィンドウサイズ
                "--window-position=" + (session["left"]).ToString() + ',' + (session["top"]).ToString(), //ウィンドウ位置
                //"--headless=new", //ヘッドレスモードを有効化
                //"--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36", //ヘッドレスモードの場合はUAの指定が必要（無いとページロードのタイムアウトが発生）
                "--enable-parallel-downloading", //#並列ダウンロードを有効化
                "--enable-quic", //QUICプロトコルを有効化
                "--test-type=gpu", //アドレスバー下に表示される「Chrome for Testing...」を非表示
                "--hide-scrollbars", //スクロールバー非表示
                "--mute-audio", //ミュート
                "--disable-background-networking", //拡張機能の更新、セーフブラウジングサービス、アップグレード検出、翻訳、UMAを含む様々なバックグラウンドネットワークサービスを無効化
                "--ignore-certificate-errors", //SSL認証(この接続ではプライバシーが保護されません)を無効化
            };
            options.AddArguments(arguments);
            options.AddExcludedArgument("enable-automation"); //「自動テストソフトウェアによって制御されています」非表示

            IWebDriver driver = new ChromeDriver(service, options);

            if ((bool)session["maximized"])
                driver.Manage().Window.Maximize();

            try
            {
                Thread.Sleep(4000); //ページ読み込み待機

                foreach (var es in new List<ReadOnlyCollection<IWebElement>> {
                    driver.FindElements(By.Id("hyv-account-frame")),
                    driver.FindElements(By.CssSelector(".mhy-dialog.mhy-interest-selector.mhy-dialog-custom.mhy-dialog-size-md.mhy-dialog-nofooter")),
                    driver.FindElements(By.ClassName("mhy-page-header-mask"))
                })
                {
                    if (es.Count > 0)
                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].remove();", es[0]);
                }
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView()", driver.FindElement(By.CssSelector(".mhy-skeleton.forum-side-skeleton.mhy-tools-side-skeleton")));
                //MessageBox.Show("正しい値を入力してください。","エラー",MessageBoxButtons.OK,MessageBoxIcon.Error);
                ClickClass(driver, "tool-logo");

                WaitVisibilityClass(driver, "list_wrap");

                //タブタイトルの変更（ページ遷移で元のhtmlのtitle属性に戻る）
                ((IJavaScriptExecutor)driver).ExecuteScript("document.title = \"Genshin Web Event Checker\"");

                //要素移動
                ((IJavaScriptExecutor)driver).ExecuteScript(
                    "arguments[1].insertBefore(arguments[0], arguments[1].firstChild);",
                    driver.FindElements(By.ClassName("list_wrap")).Last(),
                    driver.FindElement(By.Id("__layout"))
                    );

                //不要要素削除
                foreach (var class_name in new string[] {
                    "hyl-group-b",
                    "title"
                })
                {
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].remove();", WaitPresenceClass(driver, class_name)[0]);
                }

                foreach (var e in driver.FindElement(By.ClassName("list_wrap")).FindElements(By.ClassName("list")))
                {
                    if (e.FindElement(By.ClassName("list-card__header")).Text == "ウィジェット")
                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].remove();", e);
                    else
                        ((IJavaScriptExecutor)driver).ExecuteScript(@"
                        let tr_ele = document.createElement('tr');
                        arguments[0].appendChild(tr_ele);
                        tr_ele.appendChild(arguments[1]);
                        ", driver.FindElement(By.ClassName("list_wrap")), e);
                }

                //CSS編集
                if (driver.FindElement(By.ClassName("list_wrap")).FindElements(By.ClassName("list")).Count > 0)
                {
                    foreach (var e in driver.FindElements(By.ClassName("list-card__cover")))
                        ((IJavaScriptExecutor)driver).ExecuteScript(@"
                        arguments[0].style.float = 'left';
                        arguments[0].style.padding = '10px';
                        ", e);
                    foreach (var e in driver.FindElements(By.ClassName("list-card__wrap")))
                        ((IJavaScriptExecutor)driver).ExecuteScript(@"
                        arguments[0].style.padding = '10px';
                        ", e);
                    foreach (var e in driver.FindElements(By.ClassName("list-card__header")))
                        ((IJavaScriptExecutor)driver).ExecuteScript(@"
                        arguments[0].style.fontWeight = 'bold';
                        ", e);
                }
                else //イベント無しならアイコン表示
                {
                    ((IJavaScriptExecutor)driver).ExecuteScript(@"
                    arguments[0].style.display = 'table';
                    ", driver.FindElement(By.Id("__nuxt")));

                    ((IJavaScriptExecutor)driver).ExecuteScript(@"
                    arguments[0].textContent = 'Webイベントが開催されていません。';
                    arguments[0].insertAdjacentHTML('afterbegin', '<br>')
                    arguments[0].insertAdjacentHTML('afterbegin', '<img src=""https://webstatic.hoyoverse.com/upload/uploadstatic/contentweb/20210104/2021010417060635199.png"" > ')
                    arguments[0].style.fontFamily = 'Segoe';
                    arguments[0].style.fontSize = '32px';
                    arguments[0].style.verticalAlign = 'middle';
                    arguments[0].style.textAlign = 'center';
                    arguments[0].style.display = 'table-cell';
                    arguments[0].style.height = '100vh';
                    arguments[0].style.width = '100vw';
                    arguments[0].style.fontWeight = 'bold';
                    ", driver.FindElement(By.Id("__layout")));
                }
                
                string current_url;
                while (true)
                {
                    Thread.Sleep(1000);
                    current_url = driver.Url;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                driver.Quit();
            }
        }

        //Class要素表示まで待機
        static ReadOnlyCollection<IWebElement> WaitVisibilityClass(IWebDriver driver, string class_name, int timeout = 5)
        {
            return new WebDriverWait(driver, new TimeSpan(0, 0, timeout)).Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.ClassName(class_name)));
        }

        //ID要素表示まで待機
        static ReadOnlyCollection<IWebElement> WaitVisibilityId(IWebDriver driver, string id_name, int timeout = 5)
        {
            return new WebDriverWait(driver, new TimeSpan(0, 0, timeout)).Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.Id(id_name)));
        }

        //Class要素存在まで待機
        static ReadOnlyCollection<IWebElement> WaitPresenceClass(IWebDriver driver, string class_name, int timeout = 5)
        {
            return new WebDriverWait(driver, new TimeSpan(0, 0, timeout)).Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.ClassName(class_name)));
        }

        //ID要素存在まで待機
        static ReadOnlyCollection<IWebElement> WaitPresenceId(IWebDriver driver, string id_name, int timeout = 5)
        {
            return new WebDriverWait(driver, new TimeSpan(0, 0, timeout)).Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.Id(id_name)));
        }

        //Class要素クリック
        static void ClickClass(IWebDriver driver, string class_name, int index = 0, int timeout = 5)
        {
            new WebDriverWait(driver, new TimeSpan(0, 0, timeout)).Until(ExpectedConditions.ElementToBeClickable(By.ClassName(class_name)));
            driver.FindElements(By.ClassName(class_name))[index].Click();
        }

        //ID要素クリック
        static void ClickId(IWebDriver driver, string id_name, int index = 0, int timeout = 5)
        {
            new WebDriverWait(driver, new TimeSpan(0, 0, timeout)).Until(ExpectedConditions.ElementToBeClickable(By.Id(id_name)));
            driver.FindElements(By.ClassName(id_name))[index].Click();
        }
    }
}
