using System;
using System.IO;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using MyLibraries.MyEthernetLib.Enums;

namespace MyLibraries.MyEthernetLib.Classes
{
    /// <summary>
    /// Веб парсер
    /// </summary>
    public class MyWebParser
    {
        #region Items
        /// <summary>
        /// Об'єкт класу IWebDriver
        /// </summary>
        private IWebDriver webDriver = null;
        /// <summary>
        /// Посилання на сайт для парсингу
        /// </summary>
        private string url = default;
        /// <summary>
        /// Шлях до драйвера Chrome
        /// </summary>
        private string pathToChromeDriver;
        /// <summary>
        /// Назва файлу драйвера Chrome
        /// </summary>
        private string nameFileChromeDriver = default;
        #endregion Items

        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="url">Посилання на сайт для парсингу</param>
        /// <param name="pathToChromeDriver">Шлях до драйвера Chrome</param>
        public MyWebParser(string url, string pathToChromeDriver)
        {
            CheckReceivedDataForCreateMyWebParser(url, pathToChromeDriver);
            SetMainItems(url, pathToChromeDriver);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="url">Посилання на сайт для парсингу</param>
        /// <param name="pathToChromeDriver">Шлях до драйвера Chrome</param>
        /// <param name="nameFileChromeDriver">Назва файлу драйвера Chrome</param>
        public MyWebParser(string url, string pathToChromeDriver, string nameFileChromeDriver)
        {
            CheckReceivedDataForCreateMyWebParser(url, pathToChromeDriver, nameFileChromeDriver, true);
            SetMainItems(url, pathToChromeDriver, nameFileChromeDriver);
        }
        /// <summary>
        /// Перевірити отримані дані для створення нового об'єкту MyWebParser. При помилці спрацьовує виключення, Exception
        /// </summary>
        /// <param name="url">Посилання на сайт для парсингу</param>
        /// <param name="pathToChromeDriver">Шлях до драйвера Chrome</param>
        /// <param name="nameFileChromeDriver">Назва файлу драйвера Chrome</param>
        /// <param name="checkNameFileChromeDriver">Чи перевіряти назву файлу драйвера Chrome</param>
        private void CheckReceivedDataForCreateMyWebParser(string url, string pathToChromeDriver, string nameFileChromeDriver = default, bool checkNameFileChromeDriver = false)
        {
            if (!Directory.Exists(pathToChromeDriver))
                throw new Exception(message: $"Некоректно надано шлях до драйвера Chrome. Наданий шлях:'{pathToChromeDriver}'");
            else if (url == default)
                throw new Exception(message: $"Надано пустий url для парсингу!");
            else if (checkNameFileChromeDriver)
            {
                if (nameFileChromeDriver == default)
                    throw new Exception(message: "Надано порожню назву файлу драйвера Chrome!");
                else if (!File.Exists($"{pathToChromeDriver}/{nameFileChromeDriver}"))
                    throw new Exception(message: $"Файл з назвою '{nameFileChromeDriver}' за шляхом '{pathToChromeDriver}' не існує!");
            }
        }
        /// <summary>
        /// Встановити основні змінні
        /// </summary>
        /// <param name="url">Посилання на сайт для парсингу</param>
        /// <param name="pathToChromeDriver">Шлях до драйвера Chrome</param>
        /// <param name="nameFileChromeDriver">Назва файлу драйвера Chrome</param>
        private void SetMainItems(string url, string pathToChromeDriver, string nameFileChromeDriver = default)
        {
            this.url = url;
            this.pathToChromeDriver = pathToChromeDriver;
            this.nameFileChromeDriver = nameFileChromeDriver;
        }
        #endregion Constructors

        #region Function
        #region Main functions
        /// <summary>
        /// Розпочати дії
        /// </summary>
        /// <param name="hideWindow">true - парсинг відбувається в режимі закритого клієнтського вікна, false - парсинг відбувається в режимі відкритого клієнтського вікна</param>
        public void Open(bool hideWindow = true)
        {
            if (webDriver != null || url == default) return;

            #region Items
            ChromeDriverService service = ChromeDriverService.CreateDefaultService(pathToChromeDriver);
            ChromeOptions options = new ChromeOptions();

            if (hideWindow)
            {
                service.HideCommandPromptWindow = true;
                options.AddArgument("headless");
            }

            webDriver = new ChromeDriver(service, options);
            #endregion Items

            webDriver.Navigate().GoToUrl(url);

            if (!hideWindow)
            {
                webDriver.Manage().Window.Maximize();

                Thread.Sleep(1000);
            }
        }
        /// <summary>
        /// Завершити дії
        /// </summary>
        public void Close()
        {
            if (webDriver == null) return;

            try { webDriver.Close(); } catch { }
        }
        /// <summary>
        /// Повернутися на задню сторінку
        /// </summary>
        public void Back()
        {
            if (webDriver == null) return;

            webDriver.Navigate().Back();
        }
        /// <summary>
        /// Повернутися на передню сторніку
        /// </summary>
        public void Forward()
        {
            if (webDriver == null) return;

            webDriver.Navigate().Forward();
        }
        /// <summary>
        /// Обновити поточну сторінку
        /// </summary>
        public void Refresh()
        {
            if (webDriver == null) return;

            webDriver.Navigate().Refresh();
        }
        #endregion Main functions

        #region Find
        /// <summary>
        /// Знайти елемент по XPath-шляху
        /// </summary>
        /// <param name="element">Поточний веб елемент</param>
        /// <param name="XPath">XPath-шлях до веб елементу</param>
        /// <param name="setFocusOnElement">Встановити фокус на веб елементі</param>
        /// <returns>true - вдалося знайти елемент, false - не вдалося знайти елемент</returns>
        public bool FindByXPath(ref IWebElement element, string XPath, bool setFocusOnElement = false)
        {
            try
            {
                element = webDriver.FindElement(By.XPath(XPath));
            }
            catch { return false; }

            if (setFocusOnElement) SetFocusOnElement(element);

            return true;
        }
        #endregion Find

        #region Gets
        /// <summary>
        /// Отримати текст веб елемента
        /// </summary>
        /// <param name="textElement">Поточний текст веб елемента</param>
        /// <param name="XPath">XPath-шлях до веб елементу</param>
        /// <param name="setFocusOnElement">Встановити фокус на веб елементі</param>
        /// <returns>true - вдалося отримати текст веб елемента, false - не вдалося отримати текст веб елемента</returns>
        public bool GetTextElement(ref string textElement, string XPath, bool setFocusOnElement = false)
        {
            IWebElement element = default; if (!FindByXPath(ref element, XPath, setFocusOnElement)) return false;

            textElement = element.Text; return true;
        }
        /// <summary>
        /// Отримати значення атрибуту веб елемента
        /// </summary>
        /// <param name="attributeValue">Значення атрибуту</param>
        /// <param name="XPath">XPath-шлях до елемента</param>
        /// <param name="attribute">Атрибут веб елемента</param>
        /// <returns>true - вдалося отримати значення атрибуту, false - не вдалося отримати значення атрибуту</returns>
        public bool GetElementAttributeValue(ref string attributeValue, string XPath, AttributesWebElements attribute)
        {
            return GetElementAttributeValue(ref attributeValue, XPath, attribute.ToString().ToLower());
        }
        /// <summary>
        /// Отримати значення атрибуту веб елемента
        /// </summary>
        /// <param name="attributeValue">Значення атрибуту</param>
        /// <param name="XPath">XPath-шлях до елемента</param>
        /// <param name="nameAttribute">Назва атрибуту веб елемента</param>
        /// <returns>true - вдалося отримати значення атрибуту, false - не вдалося отримати значення атрибуту</returns>
        public bool GetElementAttributeValue(ref string attributeValue, string XPath, string nameAttribute)
        {
            IWebElement element = default; if (!FindByXPath(ref element, XPath)) return false;

            attributeValue = element.GetAttribute(nameAttribute);

            return true;
        }
        #endregion Gets

        #region Sets
        /// <summary>
        /// Встановити посилання на сайт для парсингу
        /// </summary>
        /// <param name="newUrl">Нове посилання на сайт для парсингу</param>
        public void SetUrl(string newUrl)
        {
            url = newUrl;

            webDriver.Navigate().GoToUrl(url);
        }
        /// <summary>
        /// Встановити фокус на веб елемент
        /// </summary>
        /// <param name="element">Веб елемент</param>
        /// <returns>true - вдалося встановити фокус на веб елемнті, false - не вдалося встановити фокус на веб елемнті</returns>
        public bool SetFocusOnElement(IWebElement element)
        {
            try
            {
                Actions actions = new Actions(webDriver);
                actions.MoveToElement(element);
                actions.Perform();

                Thread.Sleep(200);

                return true;
            }
            catch { return false; }
        }
        /// <summary>
        /// Встановити фокус на веб елемент
        /// </summary>
        /// <param name="XPath">XPath-шлях до веб елементу</param>
        /// <returns>true - вдалося встановити фокус на веб елемнті, false - не вдалося встановити фокус на веб елемнті</returns>
        public bool SetFocusOnElement(string XPath)
        {
            IWebElement element = default;

            return FindByXPath(ref element, XPath, true);
        }
        #endregion Sets
        #endregion Function
    }
}
