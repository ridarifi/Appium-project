using System;
using System.Drawing;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Remote;
using System.Data.SqlClient;
using System.Collections.Generic;

namespace mecatest
{
    [TestClass]
    public class UnitTest1
    {

        private const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        private const string Meca = "C:\\Program Files\\EBP\\Automotive1.1FRFR30\\EBP.Automotive.Application.exe";
        private const string connectionString = "Server=DESKTOP-AV3FS03\\REDASQL;uid=sa;Password=@ebp78EBP;Database=meca_965dcef5-d30b-4ea9-885d-1a5a2439af63";
        protected static WindowsDriver<WindowsElement> session;
        protected static RemoteTouchScreen touchScreen;


        
        [TestMethod]
        public void AjouModifSuppCivi()
        {
            if (session == null)
            {
                Thread.Sleep(3000);
                DesiredCapabilities appCapabilities = new DesiredCapabilities();
                appCapabilities.SetCapability("app", Meca);
                appCapabilities.SetCapability("deviceName", "WindowsPC");
                session = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), appCapabilities);
                Assert.IsNotNull(session);
                Thread.Sleep(10000);
                session.SwitchTo().Window(session.WindowHandles[0]);
                session.Keyboard.SendKeys(Keys.Command + Keys.ArrowUp + Keys.Command);

                Console.WriteLine("LeftClick on \"Paramètres\" at (31,6)");
                string parametre = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"HostForm\"][@Name=\"EBP MéCa MRoad Evolution 2019 (OL Technology)\"]/Group[@Name=\"Arrimer en haut\"][starts-with(@ClassName,\"WindowsForms10\")]/MenuBar[@Name=\"Menu principal\"][starts-with(@ClassName,\"WindowsForms10\")]/MenuItem[@Name=\"Paramètres\"]";
                var winElem1 = session.FindElementByXPath(parametre);
                if (winElem1 != null)
                {
                    winElem1.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {parametre}");
                    return;
                }


                // LeftClick on "Divers" at (49,6)
                Console.WriteLine("LeftClick on \"Divers\" at (49,6)");
                string Divers = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Pane[starts-with(@ClassName,\"WindowsForms10\")]/Menu[starts-with(@ClassName,\"WindowsForms10\")]/MenuItem[@Name=\"Divers\"]";
                var winElem2 = session.FindElementByXPath(Divers);
                if (winElem2 != null)
                {
                    winElem2.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {Divers}");
                    return;
                }


                // LeftClick on "Civilités" at (68,15)
                Console.WriteLine("LeftClick on \"Civilités\" at (68,15)");
                string Civilités = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Pane[starts-with(@ClassName,\"WindowsForms10\")]/Menu[starts-with(@ClassName,\"WindowsForms10\")]/MenuItem[@Name=\"Civilités\"]";
                var winElem3 = session.FindElementByXPath(Civilités);
                if (winElem3 != null)
                {
                    winElem3.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {Civilités}");
                    return;
                }
                session.SwitchTo().Window(session.WindowHandles[0]);

                // LeftClick on "Ajouter" at (39,16)
                Console.WriteLine("LeftClick on \"Ajouter\" at (39,16)");
                string Ajouter = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"HostForm\"][@Name=\"EBP MéCa MRoad Evolution 2019 (OL Technology)\"]/Pane[@AutomationId=\"ListPage\"]/Group[@Name=\"Arrimer en haut\"][starts-with(@ClassName,\"WindowsForms10\")]/ToolBar[@Name=\"1 personnalisé\"][starts-with(@ClassName,\"WindowsForms10\")]/Button[@Name=\"Ajouter\"]";
                var winElem5 = session.FindElementByXPath(Ajouter);
                if (winElem5 != null)
                {
                    winElem5.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {Ajouter}");
                    return;
                }
                Thread.Sleep(4000);

                // LeftClick on "526442" at (144,8)
                Console.WriteLine("LeftClick on \"526442\" at (144,8)");
                string xp6 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"CivilityEntryForm\"][@Name=\"Civilité (Nouveau)\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Edit[@AutomationId=\"CaptionTextEditor\"]/Edit[starts-with(@ClassName,\"WindowsForms10\")]";
                var winElem6 = session.FindElementByXPath(xp6);

                // KeyboardInput VirtualKeys=""m"" CapsLock=False NumLock=True ScrollLock=False
                Console.WriteLine("KeyboardInput VirtualKeys=\"\"m\"\" CapsLock=False NumLock=True ScrollLock=False");
                winElem6.SendKeys("test");
                session.SwitchTo().Window(session.WindowHandles[0]);
                Thread.Sleep(4000);

                // LeftClick on "Enregistrer et Fermer" at (77,16)
                string enregistrer = "Enregistrer et Fermer";
                var winElem7 = session.FindElementByName(enregistrer);
                winElem7.Click();

                session.SwitchTo().Window(session.WindowHandles[0]);
                Thread.Sleep(6000);
                session.Keyboard.SendKeys(Keys.Command + Keys.ArrowUp + Keys.Command);
                session.FindElementByName("Sélection row 0").Click();
                session.FindElementByName("Sélection row 20").Click();
                session.FindElementByName("Modifier").Click();

                session.SwitchTo().Window(session.WindowHandles[0]);


                // LeftClick on "test" at (104,9)
                string xp8 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"CivilityEntryForm\"][@Name=\"Civilité test\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Edit[@AutomationId=\"CaptionTextEditor\"][@Name=\"test\"]/Edit[@Name=\"test\"][starts-with(@ClassName,\"WindowsForms10\")]";
                var winElem8 = session.FindElementByXPath(xp8);
                if (winElem8 != null)
                {
                    winElem8.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {xp8}");
                    return;
                }

                winElem8.SendKeys("test");
                winElem8.SendKeys(Keys.NumberPad2);
                Thread.Sleep(4000);
                string enregistrerr = "Enregistrer et Fermer";
                var winElem9 = session.FindElementByName(enregistrerr);
                winElem9.Click();

                session.SwitchTo().Window(session.WindowHandles[0]);
                // LeftClick on "&Non" at (64,21)
                session.FindElementByAccessibilityId("bt2").Click();
                /////////////Test query////////////////
                CreerCommande(connectionString, "select top 1 Caption from Civility");
                ///////////////////////////////////////
                Thread.Sleep(4000);
                session.SwitchTo().Window(session.WindowHandles[0]);
                session.FindElementByName("Supprimer").Click();
                session.SwitchTo().Window(session.WindowHandles[0]);
                session.FindElementByAccessibilityId("bt2").Click();
                session.Quit();

            }
        }

        [TestMethod]
        public void AjouterDevisExistClient()
        {
            if (session == null)
            {
                Thread.Sleep(3000);
                DesiredCapabilities appCapabilities = new DesiredCapabilities();
                appCapabilities.SetCapability("app", Meca);
                appCapabilities.SetCapability("deviceName", "WindowsPC");
                session = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), appCapabilities);
                Assert.IsNotNull(session);
                Thread.Sleep(10000);
                session.SwitchTo().Window(session.WindowHandles[0]);

                // LeftClick on "quotesTileButton" at (124,67)
                Console.WriteLine("LeftClick on \"quotesTileButton\" at (124,67)");
                string xp1 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"HostForm\"][@Name=\"EBP MéCa MRoad Evolution 2019 (OL Technology)\"]/Pane[@AutomationId=\"WorkshopDashboardPage\"]/Pane[starts-with(@ClassName,\"WindowsForms10\")]/Pane[starts-with(@ClassName,\"WindowsForms10\")]/Pane[@AutomationId=\"DashboardTileButton\"]/Button[@AutomationId=\"quotesTileButton\"]";
                var winElem1 = session.FindElementByXPath(xp1);
                if (winElem1 != null)
                {
                    winElem1.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {xp1}");
                    return;
                }

                // LeftClick on "Ajouter" at (52,16)
                Console.WriteLine("LeftClick on \"Ajouter\" at (52,16)");
                string xp2 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"HostForm\"][@Name=\"EBP MéCa MRoad Evolution 2019 (OL Technology)\"]/Pane[@AutomationId=\"ListPage\"]/Group[@Name=\"Arrimer en haut\"][starts-with(@ClassName,\"WindowsForms10\")]/ToolBar[@Name=\"1 personnalisé\"][starts-with(@ClassName,\"WindowsForms10\")]/MenuItem[@Name=\"Ajouter\"]";
                var winElem2 = session.FindElementByXPath(xp2);
                if (winElem2 != null)
                {
                    winElem2.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {xp2}");
                    return;
                }


                // LeftClick on "Devis" at (51,13)
                Console.WriteLine("LeftClick on \"Devis\" at (51,13)");
                string xp3 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Pane[starts-with(@ClassName,\"WindowsForms10\")]/Menu[starts-with(@ClassName,\"WindowsForms10\")]/MenuItem[@Name=\"Devis\"]";
                var winElem3 = session.FindElementByXPath(xp3);
                if (winElem3 != null)
                {
                    winElem3.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {xp3}");
                    return;
                }

                session.SwitchTo().Window(session.WindowHandles[0]);
                Thread.Sleep(2000);

                // LeftClick on "Sélection row 4" at (12,8)
                Console.WriteLine("LeftClick on \"Sélection row 4\" at (12,8)");
                string xp5 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"TradeDocumentEntryFormBase\"][@Name=\"Devis du 08/04/2019 (Nouveau)\"]/Window[starts-with(@ClassName,\"WindowsForms10\")]/Pane[starts-with(@ClassName,\"WindowsForms10\")]/Pane[@AutomationId=\"list\"]/Table[@AutomationId=\"dataGrid\"]/Custom[@Name=\"Panneau de données\"]/Custom[@Name=\"Ligne 5\"]/DataItem[@Name=\"Sélection row 4\"]";
                var winElem5 = session.FindElementByXPath(xp5);
                if (winElem5 != null)
                {
                    winElem5.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {xp5}");
                    return;
                }

            }
        }

        [TestMethod]
        public void AjouterDevis()
        {
            if (session == null)
            {
                Thread.Sleep(3000);
                DesiredCapabilities appCapabilities = new DesiredCapabilities();
                appCapabilities.SetCapability("app", Meca);
                appCapabilities.SetCapability("deviceName", "WindowsPC");
                session = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), appCapabilities);
                Assert.IsNotNull(session);
                Thread.Sleep(10000);
                session.SwitchTo().Window(session.WindowHandles[0]);

                // LeftClick on "quotesTileButton" at (124,67)
                string quotesTileButton = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"HostForm\"][@Name=\"EBP MéCa MRoad Evolution 2019 (OL Technology)\"]/Pane[@AutomationId=\"WorkshopDashboardPage\"]/Pane[starts-with(@ClassName,\"WindowsForms10\")]/Pane[starts-with(@ClassName,\"WindowsForms10\")]/Pane[@AutomationId=\"DashboardTileButton\"]/Button[@AutomationId=\"quotesTileButton\"]";
                var winElemquotesTileButton = session.FindElementByXPath(quotesTileButton);
                if (winElemquotesTileButton != null)
                {
                    winElemquotesTileButton.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {quotesTileButton}");
                    return;
                }


                // LeftClick on "Ajouter" at (52,16)
                string Ajouter = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"HostForm\"][@Name=\"EBP MéCa MRoad Evolution 2019 (OL Technology)\"]/Pane[@AutomationId=\"ListPage\"]/Group[@Name=\"Arrimer en haut\"][starts-with(@ClassName,\"WindowsForms10\")]/ToolBar[@Name=\"1 personnalisé\"][starts-with(@ClassName,\"WindowsForms10\")]/MenuItem[@Name=\"Ajouter\"]";
                var winElemAjouter = session.FindElementByXPath(Ajouter);
                if (winElemAjouter != null)
                {
                    winElemAjouter.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {Ajouter}");
                    return;
                }


                // LeftClick on "Devis" at (51,13)
                string Devis = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Pane[starts-with(@ClassName,\"WindowsForms10\")]/Menu[starts-with(@ClassName,\"WindowsForms10\")]/MenuItem[@Name=\"Devis\"]";
                var winElemDevis = session.FindElementByXPath(Devis);
                if (winElemDevis != null)
                {
                    winElemDevis.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {Devis}");
                    return;
                }

                session.SwitchTo().Window(session.WindowHandles[0]);

                // LeftClick on "4852100" at (48,6) 
                string Nom = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"TradeDocumentEntryFormBase\"][@Name=\"Devis du 10/04/2019 (Nouveau)\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Edit[@AutomationId=\"thirdNameStringLookupEditor\"]/Edit[starts-with(@ClassName,\"WindowsForms10\")]";
                var winElemNom = session.FindElementByXPath(Nom);
                winElemNom.SendKeys("laura");

                // LeftClick on "1639462" at (36,9)
                string ville = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"TradeDocumentEntryFormBase\"][@Name=\"Devis du 10/04/2019 (Nouveau)\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Edit[@AutomationId=\"invoicingAddress11TextEditor\"]/Edit[starts-with(@ClassName,\"WindowsForms10\")]";
                var winElemville = session.FindElementByXPath(ville);
                winElemville.SendKeys("fqnce");

                // LeftClick on "8522258" at (55,1)
                string Adresse = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"TradeDocumentEntryFormBase\"][@Name=\"Devis du 10/04/2019 (Nouveau)\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Edit[@AutomationId=\"invoicingCity1StringLookupEditor\"]/Edit[starts-with(@ClassName,\"WindowsForms10\")]";
                var winElemAdresse = session.FindElementByXPath(Adresse);
                winElemAdresse.SendKeys("BOURG EN BRESSE");

                // LeftClick on "5442608" at (32,2)               
                string CodePostale = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"TradeDocumentEntryFormBase\"][@Name=\"Devis du 10/04/2019 (Nouveau)\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Edit[@AutomationId=\"invoicingZipCode1StringLookupEditor\"]/Edit[starts-with(@ClassName,\"WindowsForms10\")]";
                var winElemCodePostale = session.FindElementByXPath(CodePostale);
                winElemCodePostale.SendKeys(Keys.NumberPad0);
                winElemCodePostale.SendKeys(Keys.NumberPad1);
                winElemCodePostale.SendKeys(Keys.NumberPad0);
                winElemCodePostale.SendKeys(Keys.NumberPad0);
                winElemCodePostale.SendKeys(Keys.NumberPad0);



                // LeftClick on "3737666" at (52,10)
                string mail = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"TradeDocumentEntryFormBase\"][@Name=\"Devis du 10/04/2019 (Nouveau)\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Edit[@AutomationId=\"emailLinkLabelEditor\"]/Edit[starts-with(@ClassName,\"WindowsForms10\")]";
                var winElemmail = session.FindElementByXPath(mail);
                winElemmail.SendKeys("reda@gmail.com");


                // LeftClick on "vehicleRegistrationNumberLookupEditor" at (91,23)
                string Immatriculation = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"TradeDocumentEntryFormBase\"][@Name=\"Devis du 10/04/2019 (Nouveau)\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Edit[@AutomationId=\"vehicleRegistrationNumberLookupEditor\"]";
                var winElemImmatriculation = session.FindElementByXPath(Immatriculation);
                winElemImmatriculation.SendKeys(Keys.NumberPad3);
                winElemImmatriculation.SendKeys(Keys.NumberPad2);


                // LeftClick on "395984" at (14,10)               
                string xp19 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"TradeDocumentEntryFormBase\"][@Name=\"Devis du 10/04/2019 (Nouveau)\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Edit[@AutomationId=\"vehicleCountryLookupEditor\"]/Edit[starts-with(@ClassName,\"WindowsForms10\")]";
                var winElem19 = session.FindElementByXPath(xp19);
                winElem19.SendKeys("FR");


                // LeftClick on "265436" at (65,7)
                string xp15 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"TradeDocumentEntryFormBase\"][@Name=\"Devis du 10/04/2019 (Nouveau)\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Edit[@AutomationId=\"vehicleBrandLookupEditor\"]/Edit[starts-with(@ClassName,\"WindowsForms10\")]";
                var winElem15 = session.FindElementByXPath(xp15);
                winElem15.SendKeys("AIXAM");

                // LeftClick on "5833816" at (29,9)
                string xp16 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"TradeDocumentEntryFormBase\"][@Name=\"Devis du 10/04/2019 (Nouveau)\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Edit[@AutomationId=\"vehicleModelStringLookupEditor\"]/Edit[starts-with(@ClassName,\"WindowsForms10\")]";
                var winElem16 = session.FindElementByXPath(xp16);
                winElem16.SendKeys(Keys.NumberPad2);
                winElem16.SendKeys(Keys.NumberPad0);
                winElem16.SendKeys(Keys.NumberPad1);
                winElem16.SendKeys(Keys.NumberPad6);

                // LeftClick on "3934802" at (55,7)              
                string xp17 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"TradeDocumentEntryFormBase\"][@Name=\"Devis du 10/04/2019 (Nouveau)\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Edit[@AutomationId=\"vehicleVinLookupEditor\"]/Edit[starts-with(@ClassName,\"WindowsForms10\")]";
                var winElem17 = session.FindElementByXPath(xp17);
                winElem17.SendKeys(Keys.NumberPad1);
                winElem17.SendKeys(Keys.NumberPad7);

                string xp18 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"TradeDocumentEntryFormBase\"][@Name=\"Devis du 10/04/2019 (Nouveau)\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Edit[@AutomationId=\"vehicleMileageNumberEditor\"]/Edit[starts-with(@ClassName,\"WindowsForms10\")]";
                var winElem18 = session.FindElementByXPath(xp18);
                winElem18.SendKeys(Keys.NumberPad1);
                winElem18.SendKeys(Keys.NumberPad1);
                winElem18.SendKeys(Keys.NumberPad1);
                winElem18.SendKeys(Keys.NumberPad1);

                // LeftClick on "Code référence row 0" at (63,13)
                Console.WriteLine("LeftClick on \"Code référence row 0\" at (63,13)");
                string xp1 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"TradeDocumentEntryFormBase\"][@Name=\"Devis du 10/04/2019 (Nouveau)\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Tree[@AutomationId=\"linesTreeList\"]/Group[@Name=\"Panneau de données\"]/TreeItem[@Name=\"Nœud0\"]/DataItem[@Name=\"Code référence row 0\"]";
                var winElem1 = session.FindElementByXPath(xp1);
                if (winElem1 != null)
                {
                    winElem1.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {xp1}");
                    return;
                }


                // LeftClick on "Description row 0" at (76,11)
                Console.WriteLine("LeftClick on \"Description row 0\" at (76,11)");
                string xp2 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"TradeDocumentEntryFormBase\"][@Name=\"Devis du 10/04/2019 (Nouveau)\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Tree[@AutomationId=\"linesTreeList\"]/Group[@Name=\"Panneau de données\"]/TreeItem[@Name=\"Nœud0\"]/DataItem[@Name=\"Description row 0\"]";
                var winElem2 = session.FindElementByXPath(xp2);
                if (winElem2 != null)
                {
                    winElem2.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {xp2}");
                    return;
                }


                // MouseHover on "Code unité row 5" at (67,2)
                Console.WriteLine("MouseHover on \"Code unité row 5\" at (67,2)");
                string xp3 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"TradeDocumentEntryFormBase\"][@Name=\"Devis du 10/04/2019 (Nouveau)\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Tree[@AutomationId=\"linesTreeList\"]/Group[@Name=\"Panneau de données\"]/TreeItem[@Name=\"Nœud5\"]/DataItem[@Name=\"Code unité row 5\"]";
                var winElem3 = session.FindElementByXPath(xp3);
                if (winElem3 != null)
                {
                    //TODO: Hover at (67,2) on winElem3
                }
                else
                {
                    Console.WriteLine($"Failed to find element {xp3}");
                    return;
                }


                // LeftClick on "OK" at (25,10)
                Console.WriteLine("LeftClick on \"OK\" at (25,10)");
                string xp4 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"TradeDocumentEntryFormBase\"][@Name=\"Devis du 10/04/2019 (Nouveau)\"]/Window[starts-with(@ClassName,\"WindowsForms10\")]/Button[@Name=\"OK\"][starts-with(@ClassName,\"WindowsForms10\")]";
                var winElem4 = session.FindElementByXPath(xp4);
                if (winElem4 != null)
                {
                    winElem4.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {xp4}");
                    return;
                }


                // LeftClick on "Quantité row 0" at (50,10)
                Console.WriteLine("LeftClick on \"Quantité row 0\" at (50,10)");
                string xp5 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"TradeDocumentEntryFormBase\"][@Name=\"Devis du 10/04/2019 (Nouveau)\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Tree[@AutomationId=\"linesTreeList\"]/Group[@Name=\"Panneau de données\"]/TreeItem[@Name=\"Nœud0\"]/DataItem[@Name=\"Quantité row 0\"]";
                var winElem5 = session.FindElementByXPath(xp5);
                if (winElem5 != null)
                {
                    winElem5.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {xp5}");
                    return;
                }


                // LeftClick on "Code unité row 0" at (36,14)
                Console.WriteLine("LeftClick on \"Code unité row 0\" at (36,14)");
                string xp6 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"TradeDocumentEntryFormBase\"][@Name=\"Devis du 10/04/2019 (Nouveau)\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Tree[@AutomationId=\"linesTreeList\"]/Group[@Name=\"Panneau de données\"]/TreeItem[@Name=\"Nœud0\"]/DataItem[@Name=\"Code unité row 0\"]";
                var winElem6 = session.FindElementByXPath(xp6);
                if (winElem6 != null)
                {
                    winElem6.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {xp6}");
                    return;
                }




                string enregistrerr = "Enregistrer et Fermer";
                var winElemenregistrerr = session.FindElementByName(enregistrerr);
                winElemenregistrerr.Click();

                //session.Keyboard.SendKeys(Keys.Control + "e" + Keys.Control);
                session.Quit();


            }
        }

        [TestMethod]
        public void TestMethod3()
        {
            if (session == null)
            {
                Thread.Sleep(3000);
                DesiredCapabilities appCapabilities = new DesiredCapabilities();
                appCapabilities.SetCapability("app", Meca);
                appCapabilities.SetCapability("deviceName", "WindowsPC");
                session = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), appCapabilities);
                Assert.IsNotNull(session);
                Thread.Sleep(10000);
                session.SwitchTo().Window(session.WindowHandles[0]);
                session.FindElementByAccessibilityId("invoiceTileButton").Click();
                Thread.Sleep(4000);
                session.FindElementByName("Cliquer ici pour ajouter un nouvel enregistrement").Click();


                /****session.SwitchTo().Window(session.WindowHandles[0]);
                session.FindElementByName("Paramètres").Click();
                WindowsElement appNameTitle = session.FindElementByName("Divers");
                const int offset = 20;
                session.Mouse.MouseMove(appNameTitle.Coordinates);
                session.Mouse.MouseDown(null); // Pass null as this command omit the given parameter
                session.Mouse.MouseMove(appNameTitle.Coordinates, offset, offset);
                session.Mouse.MouseUp(null);***/



                /***session.SwitchTo().Window(session.WindowHandles[0]);
                session.FindElementByAccessibilityId("quotesTileButton").Click();
                Thread.Sleep(10000);
                session.FindElementByName("Cliquer ici pour ajouter un nouvel enregistrement").Click();
                Thread.Sleep(6000);
                session.SwitchTo().Window(session.WindowHandles[0]);
                session.FindElementByName("Fermer").Click();***/



            }
        }

        //flight application 
        //methode 1
        [TestMethod]
        public void TestMethod4()
        {
            if (session == null)
            {
                DesiredCapabilities appCapabilities = new DesiredCapabilities();
                appCapabilities.SetCapability("app", Meca);
                appCapabilities.SetCapability("deviceName", "WindowsPC");
                session = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), appCapabilities);
                Assert.IsNotNull(session);
                Thread.Sleep(3000);
                session.FindElementByAccessibilityId("agentName").SendKeys("john");
                Thread.Sleep(3000);
                session.FindElementByAccessibilityId("password").SendKeys("HP");
                Thread.Sleep(3000);
                session.FindElementByAccessibilityId("okButton").Click();
                Thread.Sleep(3000);

                session.SwitchTo().Window(session.WindowHandles[0]);

                session.FindElementByAccessibilityId("fromCity").Click();
                Thread.Sleep(3000);
                session.FindElementByName("Paris").Click();
                Thread.Sleep(3000);
                session.FindElementByAccessibilityId("toCity").Click();
                Thread.Sleep(3000);
                session.FindElementByName("London").Click();
                Thread.Sleep(3000);
                session.FindElementByAccessibilityId("Class").Click();
                Thread.Sleep(3000);
                session.FindElementByName("First").Click();
                Thread.Sleep(3000);
                session.FindElementByAccessibilityId("numOfTickets").Click();
                Thread.Sleep(3000);
                session.FindElementByName("5").Click();
                Thread.Sleep(3000);
                session.FindElementByName("FIND FLIGHTS").Click();

            }
        }

        [TestMethod]
        public void SupprimerDevis()
        {
            if (session == null)
            {
                Thread.Sleep(3000);
                DesiredCapabilities appCapabilities = new DesiredCapabilities();
                appCapabilities.SetCapability("app", Meca);
                appCapabilities.SetCapability("deviceName", "WindowsPC");
                session = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), appCapabilities);
                Assert.IsNotNull(session);
                Thread.Sleep(10000);
                session.SwitchTo().Window(session.WindowHandles[0]);

                // LeftClick on "quotesTileButton" at (84,61)
                string xp1 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"HostForm\"][@Name=\"EBP MéCa MRoad Evolution 2019 (OL Technology)\"]/Pane[@AutomationId=\"WorkshopDashboardPage\"]/Pane[starts-with(@ClassName,\"WindowsForms10\")]/Pane[starts-with(@ClassName,\"WindowsForms10\")]/Pane[@AutomationId=\"DashboardTileButton\"]/Button[@AutomationId=\"quotesTileButton\"]";
                var winElem1 = session.FindElementByXPath(xp1);
                if (winElem1 != null)
                {
                    winElem1.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {xp1}");
                    return;
                }

                // LeftClick on "Sélection row 0" at (14,7)
                string xp3 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"HostForm\"][@Name=\"EBP MéCa MRoad Evolution 2019 (OL Technology)\"]/Pane[@AutomationId=\"ListPage\"]/Pane[@AutomationId=\"baseList\"]/Table[@AutomationId=\"dataGrid\"][@Name=\"Cliquer ici pour ajouter un nouvel enregistrement\"]/Custom[@Name=\"Panneau de données\"]/Custom[@Name=\"Ligne 1\"]/DataItem[@Name=\"Sélection row 0\"]";
                var winElem3 = session.FindElementByXPath(xp3);
                if (winElem3 != null)
                {
                    winElem3.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {xp3}");
                    return;
                }


                // LeftClick on "Sélection row 1" at (11,12)
                string xp4 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"HostForm\"][@Name=\"EBP MéCa MRoad Evolution 2019 (OL Technology)\"]/Pane[@AutomationId=\"ListPage\"]/Pane[@AutomationId=\"baseList\"]/Table[@AutomationId=\"dataGrid\"][@Name=\"Cliquer ici pour ajouter un nouvel enregistrement\"]/Custom[@Name=\"Panneau de données\"]/Custom[@Name=\"Ligne 2\"]/DataItem[@Name=\"Sélection row 1\"]";
                var winElem4 = session.FindElementByXPath(xp4);
                if (winElem4 != null)
                {
                    winElem4.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {xp4}");
                    return;
                }


                // LeftClick on "Supprimer" at (16,15)
                Console.WriteLine("LeftClick on \"Supprimer\" at (16,15)");
                string xp5 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"HostForm\"][@Name=\"EBP MéCa MRoad Evolution 2019 (OL Technology)\"]/Pane[@AutomationId=\"ListPage\"]/Group[@Name=\"Arrimer en haut\"][starts-with(@ClassName,\"WindowsForms10\")]/ToolBar[@Name=\"1 personnalisé\"][starts-with(@ClassName,\"WindowsForms10\")]/Button[@Name=\"Supprimer\"]";
                var winElem5 = session.FindElementByXPath(xp5);
                if (winElem5 != null)
                {
                    winElem5.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {xp5}");
                    return;
                }


                // LeftClick on "&Oui" at (34,10)
                Console.WriteLine("LeftClick on \"&Oui\" at (34,10)");
                string xp6 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"HostForm\"][@Name=\"EBP MéCa MRoad Evolution 2019 (OL Technology)\"]/Window[@AutomationId=\"FrmTaskDialog\"][@Name=\"EBP MéCa MRoad Evolution 2019 (OL Technology)\"]/Pane[@AutomationId=\"pnlButtons\"][@Name=\"pnlButtons\"]/Button[starts-with(@AutomationId,\"bt\")][@Name=\"&amp;Oui\"]";
                var winElem6 = session.FindElementByXPath(xp6);
                if (winElem6 != null)
                {
                    winElem6.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {xp6}");
                    return;
                }

            }
        }
        [TestMethod]
        public void ajoutercivilite()
        {
            if (session == null)
            {
                Thread.Sleep(3000);
                DesiredCapabilities appCapabilities = new DesiredCapabilities();
                appCapabilities.SetCapability("app", Meca);
                appCapabilities.SetCapability("deviceName", "WindowsPC");
                session = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), appCapabilities);
                Assert.IsNotNull(session);
                Thread.Sleep(10000);
                session.SwitchTo().Window(session.WindowHandles[0]);

                Console.WriteLine("LeftClick on \"Paramètres\" at (31,6)");
                string parametre = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"HostForm\"][@Name=\"EBP MéCa MRoad Evolution 2019 (OL Technology)\"]/Group[@Name=\"Arrimer en haut\"][starts-with(@ClassName,\"WindowsForms10\")]/MenuBar[@Name=\"Menu principal\"][starts-with(@ClassName,\"WindowsForms10\")]/MenuItem[@Name=\"Paramètres\"]";
                var winElem1 = session.FindElementByXPath(parametre);
                if (winElem1 != null)
                {
                    winElem1.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {parametre}");
                    return;
                }


                // LeftClick on "Divers" at (49,6)
                Console.WriteLine("LeftClick on \"Divers\" at (49,6)");
                string Divers = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Pane[starts-with(@ClassName,\"WindowsForms10\")]/Menu[starts-with(@ClassName,\"WindowsForms10\")]/MenuItem[@Name=\"Divers\"]";
                var winElem2 = session.FindElementByXPath(Divers);
                if (winElem2 != null)
                {
                    winElem2.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {Divers}");
                    return;
                }


                // LeftClick on "Civilités" at (68,15)
                Console.WriteLine("LeftClick on \"Civilités\" at (68,15)");
                string Civilités = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Pane[starts-with(@ClassName,\"WindowsForms10\")]/Menu[starts-with(@ClassName,\"WindowsForms10\")]/MenuItem[@Name=\"Civilités\"]";
                var winElem3 = session.FindElementByXPath(Civilités);
                if (winElem3 != null)
                {
                    winElem3.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {Civilités}");
                    return;
                }

                // LeftClick on "Ajouter" at (39,16)
                Console.WriteLine("LeftClick on \"Ajouter\" at (39,16)");
                string Ajouter = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"HostForm\"][@Name=\"EBP MéCa MRoad Evolution 2019 (OL Technology)\"]/Pane[@AutomationId=\"ListPage\"]/Group[@Name=\"Arrimer en haut\"][starts-with(@ClassName,\"WindowsForms10\")]/ToolBar[@Name=\"1 personnalisé\"][starts-with(@ClassName,\"WindowsForms10\")]/Button[@Name=\"Ajouter\"]";
                var winElem5 = session.FindElementByXPath(Ajouter);
                if (winElem5 != null)
                {
                    winElem5.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {Ajouter}");
                    return;
                }
                Thread.Sleep(4000);

                // LeftClick on "526442" at (144,8)
                Console.WriteLine("LeftClick on \"526442\" at (144,8)");
                string xp6 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"CivilityEntryForm\"][@Name=\"Civilité (Nouveau)\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Edit[@AutomationId=\"CaptionTextEditor\"]/Edit[starts-with(@ClassName,\"WindowsForms10\")]";
                var winElem6 = session.FindElementByXPath(xp6);

                // KeyboardInput VirtualKeys=""m"" CapsLock=False NumLock=True ScrollLock=False
                Console.WriteLine("KeyboardInput VirtualKeys=\"\"m\"\" CapsLock=False NumLock=True ScrollLock=False");
                winElem6.SendKeys("test");

                // LeftClick on "Enregistrer et Fermer" at (77,16)
                string enregistrer = "Enregistrer et Fermer";
                var winElem7 = session.FindElementByName(enregistrer);
                winElem7.Click();

                session.SwitchTo().Window(session.WindowHandles[0]);
                Thread.Sleep(6000);
                session.FindElementByName("Sélection row 0").Click();
                session.FindElementByName("Sélection row 20").Click();
                session.FindElementByName("Modifier").Click();

                session.SwitchTo().Window(session.WindowHandles[0]);


                // LeftClick on "test" at (104,9)
                Console.WriteLine("LeftClick on \"test\" at (104,9)");
                string xp8 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"CivilityEntryForm\"][@Name=\"Civilité test\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Edit[@AutomationId=\"CaptionTextEditor\"][@Name=\"test\"]/Edit[@Name=\"test\"][starts-with(@ClassName,\"WindowsForms10\")]";
                var winElem8 = session.FindElementByXPath(xp8);
                if (winElem8 != null)
                {
                    winElem8.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {xp8}");
                    return;
                }

                winElem8.SendKeys("test");
                winElem8.SendKeys(Keys.NumberPad2);
                winElem7.Click();
            }
        }

      
        [TestMethod]
        public void ModifferDevis()
        {
            if (session == null)
            {
                Thread.Sleep(3000);
                DesiredCapabilities appCapabilities = new DesiredCapabilities();
                appCapabilities.SetCapability("app", Meca);
                appCapabilities.SetCapability("deviceName", "WindowsPC");
                session = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), appCapabilities);
                Assert.IsNotNull(session);
                Thread.Sleep(10000);
                session.SwitchTo().Window(session.WindowHandles[0]);

                // LeftClick on "quotesTileButton" at (84,61)
                string xp1 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"HostForm\"][@Name=\"EBP MéCa MRoad Evolution 2019 (OL Technology)\"]/Pane[@AutomationId=\"WorkshopDashboardPage\"]/Pane[starts-with(@ClassName,\"WindowsForms10\")]/Pane[starts-with(@ClassName,\"WindowsForms10\")]/Pane[@AutomationId=\"DashboardTileButton\"]/Button[@AutomationId=\"quotesTileButton\"]";
                var winElem1 = session.FindElementByXPath(xp1);
                if (winElem1 != null)
                {
                    winElem1.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {xp1}");
                    return;
                }


                Thread.Sleep(4000);
                session.FindElementByName("Modifier").Click();
                Thread.Sleep(TimeSpan.FromSeconds(4));
                session.SwitchTo().Window(session.WindowHandles[0]);
                Thread.Sleep(TimeSpan.FromSeconds(8));

                // LeftClick on "hayat" at (76,17)
                Console.WriteLine("LeftClick on \"hayat\" at (76,17)");
                string xp2 = "/Pane[@Name=\"Bureau 1\"][@ClassName=\"#32769\"]/Window[@AutomationId=\"TradeDocumentEntryFormBase\"][@Name=\"Devis [DE00000020] du 09/04/2019\"]/Pane[@AutomationId=\"layoutManager\"][@Name=\"The XtraLayoutControl\"]/Edit[@AutomationId=\"thirdNameStringLookupEditor\"][@Name=\"hayat\"]";
                var winElem2 = session.FindElementByXPath(xp2);
                if (winElem2 != null)
                {
                    winElem2.Click();
                }
                else
                {
                    Console.WriteLine($"Failed to find element {xp2}");
                    return;
                }

                winElem2.SendKeys("laura");


            }
        }

        [TestMethod]
        public void TearDown()
        {
            if (session != null)
            {
                session.Quit();
                session = null;
            }
        }

        public List<string> CreerCommande(string pConnectionString, string pQuery)
        {
            List<string> lColumn = new List<string>();
            SqlDataReader rd = null;
            using (SqlConnection con = new SqlConnection(pConnectionString))
            {
                try
                {
                    SqlCommand lSqlCommand = new SqlCommand(pQuery, con);
                    lSqlCommand.Connection.Open();
                    rd = lSqlCommand.ExecuteReader();

                    while (rd.Read())
                    {
                        lColumn.Add((string)rd[0]);
                    }

                }
                catch (Exception)
                {

                    throw;
                }
                finally
                {
                    // close the reader
                    if (rd != null)
                    {
                        rd.Close();
                    }

                    // 5. Close the connection
                    if (con != null)
                    {
                        con.Close();
                    }


                }
                return lColumn;
            }
        }

    }
}

