using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using NTVXApi.Models;
using NTVXApi.Services;
using System.IO;
using System;
using System.Collections.Generic;
using System.Text;
using System.Web.Mvc;
using System.Xml;
using System.Web;
using System.Linq;
using System.Globalization;

namespace NTVXApi.Controllers
{
    public class WebClientController : Controller
    {
        // GET: WebClient
        [HttpPost]
        public ActionResult Process(BvxRequestModel input)
        {

            // Excel parameters
            string templateFile = "NVX_B.xlsb";
            string copySheetsAndRanges = "Inputs~A1:A4~Blank42~I15;Inputs~B1:B8~Blank42~I23;Inputs~C1:C3~Blank42~I33;Inputs~D1:D7~Blank42~M15;Inputs~E1:E4~Blank42~N15;Inputs~F1:F2~Blank42~O17;Inputs~J1:J5~Blank42~M31;Inputs~K1~Blank42~N31;Inputs~L1~Blank42~O35;Inputs~M1:M3~Blank42~I37";
            string returnSheetsAndRanges = "Valuation 1~B9:N19;Income 3~A2:N45;BalanceSheet 4~A2:N52;CF Proj 5~A2:N56;CF Stmt 6~A2:N53;ROI 7~A2:N42";
            string macroToExecute = input.ButtonName == "Deal Optimizer" ? "Module2.RunDO" : "Module2.RunDQ";
            string downloadMacro = "Module2.ExportToExcel";

            string fileName = Guid.NewGuid().ToString("N") + ".xlsx";
            string downloadServerFileName = Server.MapPath("~/Excel") + @"\" + fileName;

            // creating sheets to send to Excel API
            Console.WriteLine("I am here");
            List<WebCell> cells = new List<WebCell>();
            cells.Add(new WebCell("A1", input.sales));
            cells.Add(new WebCell("A2", input.ebitda));
            cells.Add(new WebCell("A3", input.expenses));
            cells.Add(new WebCell("A4", input.rent));
            /*cells.Add(new WebCell("A5", input.adjEbitda));
            cells.Add(new WebCell("A6", input.ebitdaMargin));
*/
            //Section 2
            cells.Add(new WebCell("B1", input.operatingCash));
            cells.Add(new WebCell("B2", input.bookValue));
            cells.Add(new WebCell("B3", input.MarketValue));
            cells.Add(new WebCell("B4", input.receivable));
            cells.Add(new WebCell("B5", input.inventory));
            cells.Add(new WebCell("B6", input.miscassets));
            cells.Add(new WebCell("B7", input.ap));
            cells.Add(new WebCell("B8", input.miscliab));

            //section 3
            cells.Add(new WebCell("C1", input.salesgrowth));
            cells.Add(new WebCell("C2", input.capebditd));
            cells.Add(new WebCell("C3", input.ebditdamargin));

            //section 4
            cells.Add(new WebCell("D1", input.revarrt));
            cells.Add(new WebCell("E1", input.revarinterest));
            cells.Add(new WebCell("D2", input.revinvrt));
            cells.Add(new WebCell("E2", input.revinvinterest));
            cells.Add(new WebCell("D3", input.loanrt));
            cells.Add(new WebCell("E3", input.loaninterest));
            cells.Add(new WebCell("F1", input.loanyear));
            cells.Add(new WebCell("D4", input.caprt));
            cells.Add(new WebCell("E4", input.capinterest));
            cells.Add(new WebCell("F2", input.capyear));
            cells.Add(new WebCell("D5", input.openrt));
            cells.Add(new WebCell("D6", input.equityrt));
            cells.Add(new WebCell("D7", input.taxrt));

            //section 5
            cells.Add(new WebCell("G1", input.stockamt));
            cells.Add(new WebCell("G2", input.purchamt));
            cells.Add(new WebCell("G3", input.selleramt));
            cells.Add(new WebCell("G4", input.compamt));
            cells.Add(new WebCell("G5", input.peramt));
            cells.Add(new WebCell("G6", input.consultingamt));
            cells.Add(new WebCell("H1", input.sellerinterest));
            cells.Add(new WebCell("I1", input.selleryear));
            cells.Add(new WebCell("H2", input.compinterest));
            cells.Add(new WebCell("I2", input.compyear));
            cells.Add(new WebCell("H3", input.perinterest));
            cells.Add(new WebCell("I3", input.peryear));
            cells.Add(new WebCell("I4", input.consultingyear));

            //section 6
            cells.Add(new WebCell("J1", input.cashreserve));
            cells.Add(new WebCell("K1", input.cashreserveinterest));
            cells.Add(new WebCell("J2", input.dividend));
            cells.Add(new WebCell("J3", input.acqexp));
            cells.Add(new WebCell("J4", input.costexit));
            cells.Add(new WebCell("J5", input.oldassets));
            cells.Add(new WebCell("L1", input.newassets));

            //section 7
            cells.Add(new WebCell("M1", input.corp));
            cells.Add(new WebCell("M2", input.fadtax));
            cells.Add(new WebCell("M3", input.statetax));

            List<WebSheet> sheets = new List<WebSheet>();
            WebSheet sheet = new WebSheet() { Name = "Inputs", Cells = cells };
            sheets.Add(sheet);


            NtvxModel model = new NtvxModel()
            {
                ClientFile = "Web",
                ClientButton = input.ButtonName,
                TemplateFile = templateFile,
                InputXmlString = GetXml(sheets),
                CopySheetsAndRanges = copySheetsAndRanges,
                ReturnSheetsAndRanges = returnSheetsAndRanges,
                MacroToExecute = macroToExecute,
                DeleteFile = true,
                DownloadMacro = downloadMacro,
                DownloadFileName = downloadServerFileName
            };

            ExcelService excel = new ExcelService(new AuthService());
            ActionResult result = excel.Process(model);

            if (result is HttpStatusCodeResult)
            {
                HttpContext.Response.Headers.Add("WWW-Authenticate", @"basic realm=""NTV Web API""");
                return result;
            }

            // parsing the xml returned by the API
            string xml = ((ContentResult)result).Content;
            sheets = ParseXml(xml);

            WebSheet valuation = sheets[0];
            WebSheet income = sheets[1];
            WebSheet balance = sheets[2];
            WebSheet cashflow = sheets[3];
            WebSheet cashstate = sheets[4];
            WebSheet roi = sheets[5];

            BvxResponseModel respone = new BvxResponseModel()
            {
                ebitdaMultiple = valuation.Cells.Find(c => c.Key == "D9").Value,
                ebitMultiple = valuation.Cells.Find(c => c.Key == "D10").Value,
                ebitda = valuation.Cells.Find(c => c.Key == "D12").Value,
                bvAsset = valuation.Cells.Find(c => c.Key == "D13").Value,
                fmvAsset = valuation.Cells.Find(c => c.Key == "D14").Value,
                goodwill = valuation.Cells.Find(c => c.Key == "D15").Value,
                stockassets = valuation.Cells.Find(c => c.Key == "D16").Value,

                ev = valuation.Cells.Find(c => c.Key == "H9").Value,
                cashclosing = valuation.Cells.Find(c => c.Key == "H12").Value,
                sellernote = valuation.Cells.Find(c => c.Key == "H13").Value,
                sellerball = valuation.Cells.Find(c => c.Key == "H14").Value,
                noncompete = valuation.Cells.Find(c => c.Key == "H15").Value,
                personalgoodwill = valuation.Cells.Find(c => c.Key == "H16").Value,
                remcons = valuation.Cells.Find(c => c.Key == "H17").Value,
                evalue = valuation.Cells.Find(c => c.Key == "H18").Value,

                cashclosingper = valuation.Cells.Find(c => c.Key == "I12").Value,
                sellernoteper = valuation.Cells.Find(c => c.Key == "I13").Value,
                sellerballper = valuation.Cells.Find(c => c.Key == "I14").Value,
                noncompeteper = valuation.Cells.Find(c => c.Key == "I15").Value,
                personalgoodwillper = valuation.Cells.Find(c => c.Key == "I16").Value,
                remconsper = valuation.Cells.Find(c => c.Key == "I17").Value,
                evalueper = valuation.Cells.Find(c => c.Key == "I18").Value,

                buyerequ = valuation.Cells.Find(c => c.Key == "M9").Value,
                buyerequper = valuation.Cells.Find(c => c.Key == "N9").Value,
                buyerroe = valuation.Cells.Find(c => c.Key == "M10").Value,
                buyerequity = valuation.Cells.Find(c => c.Key == "M12").Value,
                revolverterm = valuation.Cells.Find(c => c.Key == "M13").Value,
                overadvloan = valuation.Cells.Find(c => c.Key == "M14").Value,
                mezzfinan = valuation.Cells.Find(c => c.Key == "M15").Value,
                totalcapraised = valuation.Cells.Find(c => c.Key == "M16").Value,
                lessacquisition = valuation.Cells.Find(c => c.Key == "M17").Value,
                lessexcesscash = valuation.Cells.Find(c => c.Key == "M18").Value,
                cashtosellerclosing = valuation.Cells.Find(c => c.Key == "M19").Value,


                //Income Statement
                //};
                // return JsonContent(respone);
                /*
                            IncomeResponse ires = new IncomeResponse()
                            {*/
                sales0 = income.Cells.Find(c => c.Key == "E8").Value,
                sales1 = income.Cells.Find(c => c.Key == "F8").Value,
                sales2 = income.Cells.Find(c => c.Key == "G8").Value,
                sales3 = income.Cells.Find(c => c.Key == "H8").Value,
                sales4 = income.Cells.Find(c => c.Key == "I8").Value,
                sales5 = income.Cells.Find(c => c.Key == "J8").Value,

                growth1 = income.Cells.Find(c => c.Key == "F9").Value,
                growth2 = income.Cells.Find(c => c.Key == "G9").Value,
                growth3 = income.Cells.Find(c => c.Key == "H9").Value,
                growth4 = income.Cells.Find(c => c.Key == "I9").Value,
                growth5 = income.Cells.Find(c => c.Key == "J9").Value,

                ebitda0 = income.Cells.Find(c => c.Key == "E10").Value,
                ebitda1 = income.Cells.Find(c => c.Key == "F10").Value,
                ebitda2 = income.Cells.Find(c => c.Key == "G10").Value,
                ebitda3 = income.Cells.Find(c => c.Key == "H10").Value,
                ebitda4 = income.Cells.Find(c => c.Key == "I10").Value,
                ebitda5 = income.Cells.Find(c => c.Key == "J10").Value,

                eper0 = income.Cells.Find(c => c.Key == "E11").Value,
                eper1 = income.Cells.Find(c => c.Key == "F11").Value,
                eper2 = income.Cells.Find(c => c.Key == "G11").Value,
                eper3 = income.Cells.Find(c => c.Key == "H11").Value,
                eper4 = income.Cells.Find(c => c.Key == "I11").Value,
                eper5 = income.Cells.Find(c => c.Key == "J11").Value,

                earnout1 = income.Cells.Find(c => c.Key == "F13").Value,
                earnout2 = income.Cells.Find(c => c.Key == "G13").Value,
                earnout3 = income.Cells.Find(c => c.Key == "H13").Value,
                earnout4 = income.Cells.Find(c => c.Key == "I13").Value,
                earnout5 = income.Cells.Find(c => c.Key == "J13").Value,

                remconspay1 = income.Cells.Find(c => c.Key == "F14").Value,
                remconspay2 = income.Cells.Find(c => c.Key == "G14").Value,
                remconspay3 = income.Cells.Find(c => c.Key == "H14").Value,
                remconspay4 = income.Cells.Find(c => c.Key == "I14").Value,
                remconspay5 = income.Cells.Find(c => c.Key == "J14").Value,

                ebitdaearnout1 = income.Cells.Find(c => c.Key == "F15").Value,
                ebitdaearnout2 = income.Cells.Find(c => c.Key == "G15").Value,
                ebitdaearnout3 = income.Cells.Find(c => c.Key == "H15").Value,
                ebitdaearnout4 = income.Cells.Find(c => c.Key == "I15").Value,
                ebitdaearnout5 = income.Cells.Find(c => c.Key == "J15").Value,

                depriciation1 = income.Cells.Find(c => c.Key == "F18").Value,
                depriciation2 = income.Cells.Find(c => c.Key == "G18").Value,
                depriciation3 = income.Cells.Find(c => c.Key == "H18").Value,
                depriciation4 = income.Cells.Find(c => c.Key == "I18").Value,
                depriciation5 = income.Cells.Find(c => c.Key == "J18").Value,

                noncompamor1 = income.Cells.Find(c => c.Key == "F19").Value,
                noncompamor2 = income.Cells.Find(c => c.Key == "G19").Value,
                noncompamor3 = income.Cells.Find(c => c.Key == "H19").Value,
                noncompamor4 = income.Cells.Find(c => c.Key == "I19").Value,
                noncompamor5 = income.Cells.Find(c => c.Key == "J19").Value,

                pergoodwillamor1 = income.Cells.Find(c => c.Key == "F20").Value,
                pergoodwillamor2 = income.Cells.Find(c => c.Key == "G20").Value,
                pergoodwillamor3 = income.Cells.Find(c => c.Key == "H20").Value,
                pergoodwillamor4 = income.Cells.Find(c => c.Key == "I20").Value,
                pergoodwillamor5 = income.Cells.Find(c => c.Key == "J20").Value,

                preconsamor1 = income.Cells.Find(c => c.Key == "F21").Value,
                preconsamor2 = income.Cells.Find(c => c.Key == "G21").Value,
                preconsamor3 = income.Cells.Find(c => c.Key == "H21").Value,
                preconsamor4 = income.Cells.Find(c => c.Key == "I21").Value,
                preconsamor5 = income.Cells.Find(c => c.Key == "J21").Value,

                acqcostamort1 = income.Cells.Find(c => c.Key == "F22").Value,
                acqcostamort2 = income.Cells.Find(c => c.Key == "G22").Value,
                acqcostamort3 = income.Cells.Find(c => c.Key == "H22").Value,
                acqcostamort4 = income.Cells.Find(c => c.Key == "I22").Value,
                acqcostamort5 = income.Cells.Find(c => c.Key == "J22").Value,

                goodwillamorttax1 = income.Cells.Find(c => c.Key == "F23").Value,
                goodwillamorttax2 = income.Cells.Find(c => c.Key == "G23").Value,
                goodwillamorttax3 = income.Cells.Find(c => c.Key == "H23").Value,
                goodwillamorttax4 = income.Cells.Find(c => c.Key == "I23").Value,
                goodwillamorttax5 = income.Cells.Find(c => c.Key == "J23").Value,

                totaldepandamort1 = income.Cells.Find(c => c.Key == "F24").Value,
                totaldepandamort2 = income.Cells.Find(c => c.Key == "G24").Value,
                totaldepandamort3 = income.Cells.Find(c => c.Key == "H24").Value,
                totaldepandamort4 = income.Cells.Find(c => c.Key == "I24").Value,
                totaldepandamort5 = income.Cells.Find(c => c.Key == "J24").Value,

                ebit1 = income.Cells.Find(c => c.Key == "F26").Value,
                ebit2 = income.Cells.Find(c => c.Key == "G26").Value,
                ebit3 = income.Cells.Find(c => c.Key == "H26").Value,
                ebit4 = income.Cells.Find(c => c.Key == "I26").Value,
                ebit5 = income.Cells.Find(c => c.Key == "J26").Value,

                intexprevolver1 = income.Cells.Find(c => c.Key == "F29").Value,
                intexprevolver2 = income.Cells.Find(c => c.Key == "G29").Value,
                intexprevolver3 = income.Cells.Find(c => c.Key == "H29").Value,
                intexprevolver4 = income.Cells.Find(c => c.Key == "I29").Value,
                intexprevolver5 = income.Cells.Find(c => c.Key == "J29").Value,

                intexptermloan1 = income.Cells.Find(c => c.Key == "F30").Value,
                intexptermloan2 = income.Cells.Find(c => c.Key == "G30").Value,
                intexptermloan3 = income.Cells.Find(c => c.Key == "H30").Value,
                intexptermloan4 = income.Cells.Find(c => c.Key == "I30").Value,
                intexptermloan5 = income.Cells.Find(c => c.Key == "J30").Value,

                intexpoveradvloan1 = income.Cells.Find(c => c.Key == "F31").Value,
                intexpoveradvloan2 = income.Cells.Find(c => c.Key == "G31").Value,
                intexpoveradvloan3 = income.Cells.Find(c => c.Key == "H31").Value,
                intexpoveradvloan4 = income.Cells.Find(c => c.Key == "I31").Value,
                intexpoveradvloan5 = income.Cells.Find(c => c.Key == "J31").Value,

                intexpmezzaninefinan1 = income.Cells.Find(c => c.Key == "F32").Value,
                intexpmezzaninefinan2 = income.Cells.Find(c => c.Key == "G32").Value,
                intexpmezzaninefinan3 = income.Cells.Find(c => c.Key == "H32").Value,
                intexpmezzaninefinan4 = income.Cells.Find(c => c.Key == "I32").Value,
                intexpmezzaninefinan5 = income.Cells.Find(c => c.Key == "J32").Value,

                intexpcapexloan1 = income.Cells.Find(c => c.Key == "F33").Value,
                intexpcapexloan2 = income.Cells.Find(c => c.Key == "G33").Value,
                intexpcapexloan3 = income.Cells.Find(c => c.Key == "H33").Value,
                intexpcapexloan4 = income.Cells.Find(c => c.Key == "I33").Value,
                intexpcapexloan5 = income.Cells.Find(c => c.Key == "J33").Value,

                intexpgapnote1 = income.Cells.Find(c => c.Key == "F34").Value,
                intexpgapnote2 = income.Cells.Find(c => c.Key == "G34").Value,
                intexpgapnote3 = income.Cells.Find(c => c.Key == "H34").Value,
                intexpgapnote4 = income.Cells.Find(c => c.Key == "I34").Value,
                intexpgapnote5 = income.Cells.Find(c => c.Key == "J34").Value,

                intexpgapballoonnote1 = income.Cells.Find(c => c.Key == "F35").Value,
                intexpgapballoonnote2 = income.Cells.Find(c => c.Key == "G35").Value,
                intexpgapballoonnote3 = income.Cells.Find(c => c.Key == "H35").Value,
                intexpgapballoonnote4 = income.Cells.Find(c => c.Key == "I35").Value,
                intexpgapballoonnote5 = income.Cells.Find(c => c.Key == "J35").Value,

                intexpnoncompete1 = income.Cells.Find(c => c.Key == "F36").Value,
                intexpnoncompete2 = income.Cells.Find(c => c.Key == "G36").Value,
                intexpnoncompete3 = income.Cells.Find(c => c.Key == "H36").Value,
                intexpnoncompete4 = income.Cells.Find(c => c.Key == "I36").Value,
                intexpnoncompete5 = income.Cells.Find(c => c.Key == "J36").Value,

                intexppersonalgoodwill1 = income.Cells.Find(c => c.Key == "F37").Value,
                intexppersonalgoodwill2 = income.Cells.Find(c => c.Key == "G37").Value,
                intexppersonalgoodwill3 = income.Cells.Find(c => c.Key == "H37").Value,
                intexppersonalgoodwill4 = income.Cells.Find(c => c.Key == "I37").Value,
                intexppersonalgoodwill5 = income.Cells.Find(c => c.Key == "J37").Value,

                intincomeoncash1 = income.Cells.Find(c => c.Key == "F38").Value,
                intincomeoncash2 = income.Cells.Find(c => c.Key == "G38").Value,
                intincomeoncash3 = income.Cells.Find(c => c.Key == "H38").Value,
                intincomeoncash4 = income.Cells.Find(c => c.Key == "I38").Value,
                intincomeoncash5 = income.Cells.Find(c => c.Key == "J38").Value,

                totalIntExpense1 = income.Cells.Find(c => c.Key == "F39").Value,
                totalIntExpense2 = income.Cells.Find(c => c.Key == "G39").Value,
                totalIntExpense3 = income.Cells.Find(c => c.Key == "H39").Value,
                totalIntExpense4 = income.Cells.Find(c => c.Key == "I39").Value,
                totalIntExpense5 = income.Cells.Find(c => c.Key == "I39").Value,

                taxableincome1 = income.Cells.Find(c => c.Key == "F41").Value,
                taxableincome2 = income.Cells.Find(c => c.Key == "G41").Value,
                taxableincome3 = income.Cells.Find(c => c.Key == "H41").Value,
                taxableincome4 = income.Cells.Find(c => c.Key == "I41").Value,
                taxableincome5 = income.Cells.Find(c => c.Key == "J41").Value,

                corpTaxesState1 = income.Cells.Find(c => c.Key == "F43").Value,
                corpTaxesState2 = income.Cells.Find(c => c.Key == "G43").Value,
                corpTaxesState3 = income.Cells.Find(c => c.Key == "H43").Value,
                corpTaxesState4 = income.Cells.Find(c => c.Key == "I43").Value,
                corpTaxesState5 = income.Cells.Find(c => c.Key == "J43").Value,

                corptexfederal1 = income.Cells.Find(c => c.Key == "F44").Value,
                corptexfederal2 = income.Cells.Find(c => c.Key == "G44").Value,
                corptexfederal3 = income.Cells.Find(c => c.Key == "H44").Value,
                corptexfederal4 = income.Cells.Find(c => c.Key == "I44").Value,
                corptexfederal5 = income.Cells.Find(c => c.Key == "J44").Value,

                netincome1 = income.Cells.Find(c => c.Key == "F45").Value,
                netincome2 = income.Cells.Find(c => c.Key == "G45").Value,
                netincome3 = income.Cells.Find(c => c.Key == "H45").Value,
                netincome4 = income.Cells.Find(c => c.Key == "I45").Value,
                netincome5 = income.Cells.Find(c => c.Key == "J45").Value,

                //Balance sheet

                balancepurhaseCash = balance.Cells.Find(c => c.Key == "D8").Value,
                balanceopeningCash = balance.Cells.Find(c => c.Key == "E8").Value,
                balanceCashyear1 = balance.Cells.Find(c => c.Key == "F8").Value,
                balanceCashyear2 = balance.Cells.Find(c => c.Key == "G8").Value,
                balanceCashyear3 = balance.Cells.Find(c => c.Key == "H8").Value,
                balanceCashyear4 = balance.Cells.Find(c => c.Key == "I8").Value,
                balanceCashyear5 = balance.Cells.Find(c => c.Key == "J8").Value,

                balancepurchaseAr = balance.Cells.Find(c => c.Key == "D9").Value,
                balanceopeningAr = balance.Cells.Find(c => c.Key == "E9").Value,
                balanceAryear1 = balance.Cells.Find(c => c.Key == "F9").Value,
                balanceAryear2 = balance.Cells.Find(c => c.Key == "G9").Value,
                balanceAryear3 = balance.Cells.Find(c => c.Key == "H9").Value,
                balanceAryear4 = balance.Cells.Find(c => c.Key == "I9").Value,
                balanceAryear5 = balance.Cells.Find(c => c.Key == "J9").Value,

                balancepurchaseInventory = balance.Cells.Find(c => c.Key == "D10").Value,
                balanceopeningInventory = balance.Cells.Find(c => c.Key == "E10").Value,
                balanceInventoryyear1 = balance.Cells.Find(c => c.Key == "F10").Value,
                balanceInventoryyear2 = balance.Cells.Find(c => c.Key == "G10").Value,
                balanceInventoryyear3 = balance.Cells.Find(c => c.Key == "H10").Value,
                balanceInventoryyear4 = balance.Cells.Find(c => c.Key == "I10").Value,
                balanceInventoryyear5 = balance.Cells.Find(c => c.Key == "J10").Value,


                balancepurchasemiscAssets = balance.Cells.Find(c => c.Key == "D11").Value,
                balanceopeningmiscAssets = balance.Cells.Find(c => c.Key == "E11").Value,
                balancemiscAssetsyear1 = balance.Cells.Find(c => c.Key == "F11").Value,
                balancemiscAssetsyear2 = balance.Cells.Find(c => c.Key == "G11").Value,
                balancemiscAssetsyear3 = balance.Cells.Find(c => c.Key == "H11").Value,
                balancemiscAssetsyear4 = balance.Cells.Find(c => c.Key == "I11").Value,
                balancemiscAssetsyear5 = balance.Cells.Find(c => c.Key == "J11").Value,

                balancepurchaseFixedAssOld = balance.Cells.Find(c => c.Key == "D12").Value,
                balanceopeningFixedAssOld = balance.Cells.Find(c => c.Key == "E12").Value,
                balanceFixedAssOldyear1 = balance.Cells.Find(c => c.Key == "F12").Value,
                balanceFixedAssOldyear2 = balance.Cells.Find(c => c.Key == "G12").Value,
                balanceFixedAssOldyear3 = balance.Cells.Find(c => c.Key == "H12").Value,
                balanceFixedAssOldyear4 = balance.Cells.Find(c => c.Key == "I12").Value,
                balanceFixedAssOldyear5 = balance.Cells.Find(c => c.Key == "J12").Value,

                balanceopeningADold = balance.Cells.Find(c => c.Key == "E13").Value,
                balanceADoldYear1 = balance.Cells.Find(c => c.Key == "F13").Value,
                balanceADoldYear2 = balance.Cells.Find(c => c.Key == "G13").Value,
                balanceADoldYear3 = balance.Cells.Find(c => c.Key == "H13").Value,
                balanceADoldYear4 = balance.Cells.Find(c => c.Key == "I13").Value,
                balanceADoldYear5 = balance.Cells.Find(c => c.Key == "J13").Value,

                balanceopenNewFixedAssets = balance.Cells.Find(c => c.Key == "E14").Value,
                balanceNewFixedAssYear1 = balance.Cells.Find(c => c.Key == "F14").Value,
                balanceNewFixedAssYear2 = balance.Cells.Find(c => c.Key == "G14").Value,
                balanceNewFixedAssYear3 = balance.Cells.Find(c => c.Key == "H14").Value,
                balanceNewFixedAssYear4 = balance.Cells.Find(c => c.Key == "I14").Value,
                balanceNewFixedAssYear5 = balance.Cells.Find(c => c.Key == "J14").Value,

                balanceopenADNewFixedAssets = balance.Cells.Find(c => c.Key == "E15").Value,
                balanceADNewFixedAssyear1 = balance.Cells.Find(c => c.Key == "F15").Value,
                balanceADNewFixedAssyear2 = balance.Cells.Find(c => c.Key == "G15").Value,
                balanceADNewFixedAssyear3 = balance.Cells.Find(c => c.Key == "H15").Value,
                balanceADNewFixedAssyear4 = balance.Cells.Find(c => c.Key == "I15").Value,
                balanceADNewFixedAssyear5 = balance.Cells.Find(c => c.Key == "J15").Value,

                balanceOpenAcquisitionExp = balance.Cells.Find(c => c.Key == "E16").Value,
                balanceAcquisitionYear1 = balance.Cells.Find(c => c.Key == "F16").Value,
                balanceAcquisitionYear2 = balance.Cells.Find(c => c.Key == "G16").Value,
                balanceAcquisitionYear3 = balance.Cells.Find(c => c.Key == "H16").Value,
                balanceAcquisitionYear4 = balance.Cells.Find(c => c.Key == "I16").Value,
                balanceAcquisitionYear5 = balance.Cells.Find(c => c.Key == "J16").Value,

                balanceOpenNonCompete = balance.Cells.Find(c => c.Key == "E17").Value,
                balanceNonCompeteYear1 = balance.Cells.Find(c => c.Key == "F17").Value,
                balanceNonCompeteYear2 = balance.Cells.Find(c => c.Key == "G17").Value,
                balanceNonCompeteYear3 = balance.Cells.Find(c => c.Key == "H17").Value,
                balanceNonCompeteYear4 = balance.Cells.Find(c => c.Key == "I17").Value,
                balanceNonCompeteYear5 = balance.Cells.Find(c => c.Key == "J17").Value,

                balanceopenPersonalGoodwill = balance.Cells.Find(c => c.Key == "E18").Value,
                balancePersonalGoodwillYear1 = balance.Cells.Find(c => c.Key == "F18").Value,
                balancePersonalGoodwillYear2 = balance.Cells.Find(c => c.Key == "G18").Value,
                balancePersonalGoodwillYear3 = balance.Cells.Find(c => c.Key == "H18").Value,
                balancePersonalGoodwillYear4 = balance.Cells.Find(c => c.Key == "I18").Value,
                balancePersonalGoodwillYear5 = balance.Cells.Find(c => c.Key == "J18").Value,

                balanceOpenRemCons1 = balance.Cells.Find(c => c.Key == "E19").Value,
                balanceRemConsYear1 = balance.Cells.Find(c => c.Key == "F19").Value,
                balanceRemConsYear2 = balance.Cells.Find(c => c.Key == "G19").Value,
                balanceRemConsYear3 = balance.Cells.Find(c => c.Key == "H19").Value,
                balanceRemConsYear4 = balance.Cells.Find(c => c.Key == "I19").Value,
                balanceRemConsYear5 = balance.Cells.Find(c => c.Key == "J19").Value,

                balanceOpenPreCons = balance.Cells.Find(c => c.Key == "E20").Value,
                balancePrepaidConsYear1 = balance.Cells.Find(c => c.Key == "F20").Value,
                balancePrepaidConsYear2 = balance.Cells.Find(c => c.Key == "G20").Value,
                balancePrepaidConsYear3 = balance.Cells.Find(c => c.Key == "H20").Value,
                balancePrepaidConsYear4 = balance.Cells.Find(c => c.Key == "I20").Value,
                balancePrepaidConsYear5 = balance.Cells.Find(c => c.Key == "J20").Value,

                balanceopenInvREntity = balance.Cells.Find(c => c.Key == "E21").Value,
                balanceInvREntityYear1 = balance.Cells.Find(c => c.Key == "F21").Value,
                balanceInvREntityYear2 = balance.Cells.Find(c => c.Key == "G21").Value,
                balanceInvREntityYear3 = balance.Cells.Find(c => c.Key == "H21").Value,
                balanceInvREntityYear4 = balance.Cells.Find(c => c.Key == "I21").Value,
                balanceInvREntityYear5 = balance.Cells.Find(c => c.Key == "J21").Value,

                balancepurchGoodwill = balance.Cells.Find(c => c.Key == "D22").Value,
                balanceopenGoodwillRes = balance.Cells.Find(c => c.Key == "E22").Value,
                balanceGoodwillYear1 = balance.Cells.Find(c => c.Key == "F22").Value,
                balanceGoodwillYear2 = balance.Cells.Find(c => c.Key == "G22").Value,
                balanceGoodwillYear3 = balance.Cells.Find(c => c.Key == "H22").Value,
                balanceGoodwillYear4 = balance.Cells.Find(c => c.Key == "I22").Value,
                balanceGoodwillYear5 = balance.Cells.Find(c => c.Key == "J22").Value,


                balancepurchTotalAssets = balance.Cells.Find(c => c.Key == "D23").Value,
                balanceopenTotalAssets = balance.Cells.Find(c => c.Key == "E23").Value,
                balanceTotalAssetsYear1 = balance.Cells.Find(c => c.Key == "F23").Value,
                balanceTotalAssetsYear2 = balance.Cells.Find(c => c.Key == "G23").Value,
                balanceTotalAssetsYear3 = balance.Cells.Find(c => c.Key == "H23").Value,
                balanceTotalAssetsYear4 = balance.Cells.Find(c => c.Key == "I23").Value,
                balanceTotalAssetsYear5 = balance.Cells.Find(c => c.Key == "J23").Value,

                balancepurchAPaccured = balance.Cells.Find(c => c.Key == "D26").Value,
                balanceopenARAccured = balance.Cells.Find(c => c.Key == "E26").Value,
                balanceARAccuredYear1 = balance.Cells.Find(c => c.Key == "F26").Value,
                balanceARAccuredYear2 = balance.Cells.Find(c => c.Key == "G26").Value,
                balanceARAccuredYear3 = balance.Cells.Find(c => c.Key == "H26").Value,
                balanceARAccuredYear4 = balance.Cells.Find(c => c.Key == "I26").Value,
                balanceARAccuredYear5 = balance.Cells.Find(c => c.Key == "J26").Value,


                balancepurchOtherMiscLiab = balance.Cells.Find(c => c.Key == "D27").Value,
                balanceopenOtherMiscLiab = balance.Cells.Find(c => c.Key == "E27").Value,
                balanceOtherMiscLiabYear1 = balance.Cells.Find(c => c.Key == "F27").Value,
                balanceOtherMiscLiabYear2 = balance.Cells.Find(c => c.Key == "G27").Value,
                balanceOtherMiscLiabYear3 = balance.Cells.Find(c => c.Key == "H27").Value,
                balanceOtherMiscLiabYear4 = balance.Cells.Find(c => c.Key == "I27").Value,
                balanceOtherMiscLiabYear5 = balance.Cells.Find(c => c.Key == "J27").Value,

                balancepurchNonOperLiab = balance.Cells.Find(c => c.Key == "D28").Value,
                balanceopenNonOperLiab = balance.Cells.Find(c => c.Key == "E28").Value,
                balanceNonOperLiabYear1 = balance.Cells.Find(c => c.Key == "F28").Value,
                balanceNonOperLiabYear2 = balance.Cells.Find(c => c.Key == "G28").Value,
                balanceNonOperLiabYear3 = balance.Cells.Find(c => c.Key == "H28").Value,
                balanceNonOperLiabYear4 = balance.Cells.Find(c => c.Key == "I28").Value,
                balanceNonOperLiabYear5 = balance.Cells.Find(c => c.Key == "J28").Value,

                balanceopenRevolver = balance.Cells.Find(c => c.Key == "E30").Value,
                balanceRevolverYear1 = balance.Cells.Find(c => c.Key == "F30").Value,
                balanceRevolverYear2 = balance.Cells.Find(c => c.Key == "G30").Value,
                balanceRevolverYear3 = balance.Cells.Find(c => c.Key == "H30").Value,
                balanceRevolverYear4 = balance.Cells.Find(c => c.Key == "I30").Value,
                balanceRevolverYear5 = balance.Cells.Find(c => c.Key == "J30").Value,

                balanceopenTermLoan = balance.Cells.Find(c => c.Key == "E31").Value,
                balanceTermLoanYear1 = income.Cells.Find(c => c.Key == "F31").Value,
                balanceTermLoanYear2 = income.Cells.Find(c => c.Key == "G31").Value,
                balanceTermLoanYear3 = income.Cells.Find(c => c.Key == "H31").Value,
                balanceTermLoanYear4 = income.Cells.Find(c => c.Key == "I31").Value,
                balanceTermLoanYear5 = balance.Cells.Find(c => c.Key == "J31").Value,

                balanceopenOverAdvLoan = balance.Cells.Find(c => c.Key == "E32").Value,
                balanceOverAdvLoan1 = balance.Cells.Find(c => c.Key == "F32").Value,
                balanceOverAdvLoan2 = balance.Cells.Find(c => c.Key == "G32").Value,
                balanceOverAdvLoan3 = balance.Cells.Find(c => c.Key == "H32").Value,
                balanceOverAdvLoan4 = balance.Cells.Find(c => c.Key == "I32").Value,
                balanceOverAdvLoan5 = balance.Cells.Find(c => c.Key == "J32").Value,

                balanceopenMezzFinancing = balance.Cells.Find(c => c.Key == "E33").Value,
                balanceMezzFinancingYear1 = balance.Cells.Find(c => c.Key == "F33").Value,
                balanceMezzFinancingYear2 = balance.Cells.Find(c => c.Key == "G33").Value,
                balanceMezzFinancingYear3 = balance.Cells.Find(c => c.Key == "H33").Value,
                balanceMezzFinancingYear4 = balance.Cells.Find(c => c.Key == "I33").Value,
                balanceMezzFinancingYear5 = balance.Cells.Find(c => c.Key == "J33").Value,

                balanceopenGapNote = balance.Cells.Find(c => c.Key == "E34").Value,
                balanceGapNoteYear1 = balance.Cells.Find(c => c.Key == "F34").Value,
                balanceGapNoteYear2 = balance.Cells.Find(c => c.Key == "G34").Value,
                balanceGapNoteYear3 = balance.Cells.Find(c => c.Key == "H34").Value,
                balanceGapNoteYear4 = balance.Cells.Find(c => c.Key == "I34").Value,
                balanceGapNoteYear5 = balance.Cells.Find(c => c.Key == "J34").Value,

                balanceopenGapBallonNote = balance.Cells.Find(c => c.Key == "E35").Value,
                balanceGapBallonNoteYear1 = balance.Cells.Find(c => c.Key == "F35").Value,
                balanceGapBallonNoteYear2 = balance.Cells.Find(c => c.Key == "G35").Value,
                balanceGapBallonNoteYear3 = balance.Cells.Find(c => c.Key == "H35").Value,
                balanceGapBallonNoteYear4 = balance.Cells.Find(c => c.Key == "I35").Value,
                balanceGapBallonNoteYear5 = balance.Cells.Find(c => c.Key == "J35").Value,

                balanceopenCapExLoan = balance.Cells.Find(c => c.Key == "E36").Value,
                balanceCapExLoanYear1 = balance.Cells.Find(c => c.Key == "F36").Value,
                balanceCapExLoanYear2 = balance.Cells.Find(c => c.Key == "G36").Value,
                balanceCapExLoanYear3 = balance.Cells.Find(c => c.Key == "H36").Value,
                balanceCapExLoanYear4 = balance.Cells.Find(c => c.Key == "I36").Value,
                balanceCapExLoanYear5 = balance.Cells.Find(c => c.Key == "J36").Value,

                balanceopenRemNonCompPayment = balance.Cells.Find(c => c.Key == "E37").Value,
                balanceRemNonComPaymentYear1 = balance.Cells.Find(c => c.Key == "F37").Value,
                balanceRemNonComPaymentYear2 = balance.Cells.Find(c => c.Key == "G37").Value,
                balanceRemNonComPaymentYear3 = balance.Cells.Find(c => c.Key == "H37").Value,
                balanceRemNonComPaymentYear4 = balance.Cells.Find(c => c.Key == "I37").Value,
                balanceRemNonComPaymentYear5 = balance.Cells.Find(c => c.Key == "J37").Value,

                balanceopenRemPersonalGWpay = balance.Cells.Find(c => c.Key == "E38").Value,
                balanceRemPersonalGWpayYear1 = balance.Cells.Find(c => c.Key == "F38").Value,
                balanceRemPersonalGWpayYear2 = balance.Cells.Find(c => c.Key == "G38").Value,
                balanceRemPersonalGWpayYear3 = balance.Cells.Find(c => c.Key == "H38").Value,
                balanceRemPersonalGWpayYear4 = balance.Cells.Find(c => c.Key == "I38").Value,
                balanceRemPersonalGWpayYear5 = balance.Cells.Find(c => c.Key == "J38").Value,

                balanceopenRemConsPay = balance.Cells.Find(c => c.Key == "E39").Value,
                balanceRemConssPayYear1 = balance.Cells.Find(c => c.Key == "F39").Value,
                balanceRemConssPayYear2 = balance.Cells.Find(c => c.Key == "G39").Value,
                balanceRemConssPayYear3 = balance.Cells.Find(c => c.Key == "H39").Value,
                balanceRemConssPayYear4 = balance.Cells.Find(c => c.Key == "I39").Value,
                balanceRemConssPayYear5 = balance.Cells.Find(c => c.Key == "J39").Value,

                balanceopenNonOperatingLiab = balance.Cells.Find(c => c.Key == "E40").Value,
                balanceNonOperatingLiabYear1 = balance.Cells.Find(c => c.Key == "F40").Value,
                balanceNonOperatingLiabYear2 = balance.Cells.Find(c => c.Key == "G40").Value,
                balanceNonOperatingLiabYear3 = balance.Cells.Find(c => c.Key == "H40").Value,
                balanceNonOperatingLiabYear4 = balance.Cells.Find(c => c.Key == "I40").Value,
                balanceNonOperatingLiabYear5 = balance.Cells.Find(c => c.Key == "J40").Value,

                balancepurchTotalLiab = balance.Cells.Find(c => c.Key == "D42").Value,
                balanceopenTotalLiab = balance.Cells.Find(c => c.Key == "E42").Value,
                balanceTotalLiabYear1 = balance.Cells.Find(c => c.Key == "F42").Value,
                balanceTotalLiabYear2 = balance.Cells.Find(c => c.Key == "G42").Value,
                balanceTotalLiabYear3 = balance.Cells.Find(c => c.Key == "H42").Value,
                balanceTotalLiabYear4 = balance.Cells.Find(c => c.Key == "I42").Value,
                balanceTotalLiabYear5 = balance.Cells.Find(c => c.Key == "J42").Value,

                balancepurchRetainedEarning = balance.Cells.Find(c => c.Key == "D44").Value,
                balanceopenRetaiedEarning = balance.Cells.Find(c => c.Key == "E44").Value,
                balanceRetainedEarningYr1 = balance.Cells.Find(c => c.Key == "F44").Value,
                balanceRetainedEarningYr2 = balance.Cells.Find(c => c.Key == "G44").Value,
                balanceRetainedEarningYr3 = balance.Cells.Find(c => c.Key == "H44").Value,
                balanceRetainedEarningYr4 = balance.Cells.Find(c => c.Key == "I44").Value,
                balanceRetainedEarningYr5 = balance.Cells.Find(c => c.Key == "J44").Value,

                balanceopenAddCapital = balance.Cells.Find(c => c.Key == "E45").Value,
                balanceAddCapitalYear1 = balance.Cells.Find(c => c.Key == "F45").Value,
                balanceAddCapitalYear2 = balance.Cells.Find(c => c.Key == "G45").Value,
                balanceAddCapitalYear3 = balance.Cells.Find(c => c.Key == "H45").Value,
                balanceAddCapitalYear4 = balance.Cells.Find(c => c.Key == "I45").Value,
                balanceAddCapitalYear5 = balance.Cells.Find(c => c.Key == "J45").Value,

                balanceopenDisTax = balance.Cells.Find(c => c.Key == "E46").Value,
                balanceDisTaxYear1 = balance.Cells.Find(c => c.Key == "F46").Value,
                balanceDisTaxYear2 = balance.Cells.Find(c => c.Key == "G46").Value,
                balanceDisTaxYear3 = balance.Cells.Find(c => c.Key == "H46").Value,
                balanceDisTaxYear4 = balance.Cells.Find(c => c.Key == "I46").Value,
                balanceDisTaxYear5 = balance.Cells.Find(c => c.Key == "J46").Value,

                balanceopenDividend = balance.Cells.Find(c => c.Key == "E47").Value,
                balanceDividendYear1 = balance.Cells.Find(c => c.Key == "F47").Value,
                balanceDividendYear2 = balance.Cells.Find(c => c.Key == "G47").Value,
                balanceDividendYear3 = balance.Cells.Find(c => c.Key == "H47").Value,
                balanceDividendYear4 = balance.Cells.Find(c => c.Key == "I47").Value,
                balanceDividendYear5 = balance.Cells.Find(c => c.Key == "J47").Value,

                balancepurchCommonStk = balance.Cells.Find(c => c.Key == "D48").Value,
                balanceopenCommonStk = balance.Cells.Find(c => c.Key == "E48").Value,
                balanceCommonStkYear1 = balance.Cells.Find(c => c.Key == "F48").Value,
                balanceCommonStkYear2 = balance.Cells.Find(c => c.Key == "G48").Value,
                balanceCommonStkYear3 = balance.Cells.Find(c => c.Key == "H48").Value,
                balanceCommonStkYear4 = balance.Cells.Find(c => c.Key == "I48").Value,
                balanceCommonStkYear5 = balance.Cells.Find(c => c.Key == "J48").Value,

                balancepurchTotalEquity = balance.Cells.Find(c => c.Key == "J36").Value,
                balanceopenTotalEquity = balance.Cells.Find(c => c.Key == "J36").Value,
                balanceTotalEquYear1 = balance.Cells.Find(c => c.Key == "J36").Value,
                balanceTotalEquYear2 = balance.Cells.Find(c => c.Key == "J36").Value,
                balanceTotalEquYear3 = balance.Cells.Find(c => c.Key == "J36").Value,
                balanceTotalEquYear4 = balance.Cells.Find(c => c.Key == "J36").Value,
                balanceTotalEquYear5 = balance.Cells.Find(c => c.Key == "J36").Value,

                balancepurchTotalLiabEqu = balance.Cells.Find(c => c.Key == "J36").Value,
                balanceopenTotalLiabEqu = balance.Cells.Find(c => c.Key == "J36").Value,
                balanceTotalLiabEquYear1 = balance.Cells.Find(c => c.Key == "J36").Value,
                balanceTotalLiabEquYear2 = balance.Cells.Find(c => c.Key == "J36").Value,
                balanceTotalLiabEquYear3 = balance.Cells.Find(c => c.Key == "J36").Value,
                balanceTotalLiabEquYear4 = balance.Cells.Find(c => c.Key == "J36").Value,
                balanceTotalLiabEquYear5 = balance.Cells.Find(c => c.Key == "J36").Value,

                //Free Cash Flow

                freeNetIncomeYr1 = cashflow.Cells.Find(c => c.Key == "F8").Value,
                freeNetIncomeYr2 = cashflow.Cells.Find(c => c.Key == "G8").Value,
                freeNetIncomeYr3 = cashflow.Cells.Find(c => c.Key == "H8").Value,
                freeNetIncomeYr4 = cashflow.Cells.Find(c => c.Key == "I8").Value,
                freeNetIncomeYr5 = cashflow.Cells.Find(c => c.Key == "J8").Value,

                freedeprYr1 = cashflow.Cells.Find(c => c.Key == "F9").Value,
                freedeprYr2 = cashflow.Cells.Find(c => c.Key == "G9").Value,
                freedeprYr3 = cashflow.Cells.Find(c => c.Key == "H9").Value,
                freedeprYr4 = cashflow.Cells.Find(c => c.Key == "I9").Value,
                freedeprYr5 = cashflow.Cells.Find(c => c.Key == "J9").Value,

                freeNonCompAmortYr1 = cashflow.Cells.Find(c => c.Key == "F10").Value,
                freeNonCompAmortYr2 = cashflow.Cells.Find(c => c.Key == "G10").Value,
                freeNonCompAmortYr3 = cashflow.Cells.Find(c => c.Key == "H10").Value,
                freeNonCompAmortYr4 = cashflow.Cells.Find(c => c.Key == "I10").Value,
                freeNonCompAmortYr5 = cashflow.Cells.Find(c => c.Key == "J10").Value,

                freePerGoodwillAmortYr1 = cashflow.Cells.Find(c => c.Key == "F11").Value,
                freePerGoodwillAmortYr2 = cashflow.Cells.Find(c => c.Key == "G11").Value,
                freePerGoodwillAmortYr3 = cashflow.Cells.Find(c => c.Key == "H11").Value,
                freePerGoodwillAmortYr4 = cashflow.Cells.Find(c => c.Key == "I11").Value,
                freePerGoodwillAmortYr5 = cashflow.Cells.Find(c => c.Key == "J11").Value,

                freePreConsAmortYr1 = cashflow.Cells.Find(c => c.Key == "F12").Value,
                freePreConsAmortYr2 = cashflow.Cells.Find(c => c.Key == "G12").Value,
                freePreConsAmortYr3 = cashflow.Cells.Find(c => c.Key == "H12").Value,
                freePreConsAmortYr4 = cashflow.Cells.Find(c => c.Key == "I12").Value,
                freePreConsAmortYr5 = cashflow.Cells.Find(c => c.Key == "J12").Value,

                freeAcqCostAmort1 = cashflow.Cells.Find(c => c.Key == "F13").Value,
                freeAcqCostAmort2 = cashflow.Cells.Find(c => c.Key == "G13").Value,
                freeAcqCostAmort3 = cashflow.Cells.Find(c => c.Key == "H13").Value,
                freeAcqCostAmort4 = cashflow.Cells.Find(c => c.Key == "I13").Value,
                freeAcqCostAmort5 = cashflow.Cells.Find(c => c.Key == "J13").Value,

                freeGoddwillAmortYr1 = cashflow.Cells.Find(c => c.Key == "F14").Value,
                freeGoddwillAmortYr2 = cashflow.Cells.Find(c => c.Key == "G14").Value,
                freeGoddwillAmortYr3 = cashflow.Cells.Find(c => c.Key == "H14").Value,
                freeGoddwillAmortYr4 = cashflow.Cells.Find(c => c.Key == "I14").Value,
                freeGoddwillAmortYr5 = cashflow.Cells.Find(c => c.Key == "J14").Value,

                freeWCchangeYr1 = cashflow.Cells.Find(c => c.Key == "F15").Value,
                freeWCchangeYr2 = cashflow.Cells.Find(c => c.Key == "G15").Value,
                freeWCchangeYr3 = cashflow.Cells.Find(c => c.Key == "H15").Value,
                freeWCchangeYr4 = cashflow.Cells.Find(c => c.Key == "I15").Value,
                freeWCchangeYr5 = cashflow.Cells.Find(c => c.Key == "J15").Value,

                freeRevPaydownWCYr1 = cashflow.Cells.Find(c => c.Key == "F16").Value,
                freeRevPaydownWCYr2 = cashflow.Cells.Find(c => c.Key == "G16").Value,
                freeRevPaydownWCYr3 = cashflow.Cells.Find(c => c.Key == "H16").Value,
                freeRevPaydownWCYr4 = cashflow.Cells.Find(c => c.Key == "I16").Value,
                freeRevPaydownWCYr5 = cashflow.Cells.Find(c => c.Key == "J16").Value,

                freeTermLoanPayYr1 = cashflow.Cells.Find(c => c.Key == "F17").Value,
                freeTermLoanPayYr2 = cashflow.Cells.Find(c => c.Key == "G17").Value,
                freeTermLoanPayYr3 = cashflow.Cells.Find(c => c.Key == "H17").Value,
                freeTermLoanPayYr4 = cashflow.Cells.Find(c => c.Key == "I17").Value,
                freeTermLoanPayYr5 = cashflow.Cells.Find(c => c.Key == "J17").Value,

                freeOverAdvLoanPayYr1 = cashflow.Cells.Find(c => c.Key == "F18").Value,
                freeOverAdvLoanPayYr2 = cashflow.Cells.Find(c => c.Key == "G18").Value,
                freeOverAdvLoanPayYr3 = cashflow.Cells.Find(c => c.Key == "H18").Value,
                freeOverAdvLoanPayYr4 = cashflow.Cells.Find(c => c.Key == "I18").Value,
                freeOverAdvLoanPayYr5 = cashflow.Cells.Find(c => c.Key == "J18").Value,

                freeMezzFinanPayYr1 = cashflow.Cells.Find(c => c.Key == "F19").Value,
                freeMezzFinanPayYr2 = cashflow.Cells.Find(c => c.Key == "G19").Value,
                freeMezzFinanPayYr3 = cashflow.Cells.Find(c => c.Key == "H19").Value,
                freeMezzFinanPayYr4 = cashflow.Cells.Find(c => c.Key == "I19").Value,
                freeMezzFinanPayYr5 = cashflow.Cells.Find(c => c.Key == "J19").Value,

                freeGapNotePayYr1 = cashflow.Cells.Find(c => c.Key == "F20").Value,
                freeGapNotePayYr2 = cashflow.Cells.Find(c => c.Key == "G20").Value,
                freeGapNotePayYr3 = cashflow.Cells.Find(c => c.Key == "H20").Value,
                freeGapNotePayYr4 = cashflow.Cells.Find(c => c.Key == "I20").Value,
                freeGapNotePayYr5 = cashflow.Cells.Find(c => c.Key == "J20").Value,

                freeGapBalloonNotePayYr1 = cashflow.Cells.Find(c => c.Key == "F21").Value,
                freeGapBalloonNotePayYr2 = cashflow.Cells.Find(c => c.Key == "G21").Value,
                freeGapBalloonNotePayYr3 = cashflow.Cells.Find(c => c.Key == "H21").Value,
                freeGapBalloonNotePayYr4 = cashflow.Cells.Find(c => c.Key == "I21").Value,
                freeGapBalloonNotePayYr5 = cashflow.Cells.Find(c => c.Key == "J21").Value,

                freeRemNonCompPayYr1 = cashflow.Cells.Find(c => c.Key == "F22").Value,
                freeRemNonCompPayYr2 = cashflow.Cells.Find(c => c.Key == "G22").Value,
                freeRemNonCompPayYr3 = cashflow.Cells.Find(c => c.Key == "H22").Value,
                freeRemNonCompPayYr4 = cashflow.Cells.Find(c => c.Key == "I22").Value,
                freeRemNonCompPayYr5 = cashflow.Cells.Find(c => c.Key == "J22").Value,

                freeRemPersGoodwillPayYr1 = cashflow.Cells.Find(c => c.Key == "F23").Value,
                freeRemPersGoodwillPayYr2 = cashflow.Cells.Find(c => c.Key == "G23").Value,
                freeRemPersGoodwillPayYr3 = cashflow.Cells.Find(c => c.Key == "H23").Value,
                freeRemPersGoodwillPayYr4 = cashflow.Cells.Find(c => c.Key == "I23").Value,
                freeRemPersGoodwillPayYr5 = cashflow.Cells.Find(c => c.Key == "J23").Value,

                freeCapExpYr1 = cashflow.Cells.Find(c => c.Key == "F24").Value,
                freeCapExpYr2 = cashflow.Cells.Find(c => c.Key == "G24").Value,
                freeCapExpYr3 = cashflow.Cells.Find(c => c.Key == "H24").Value,
                freeCapExpYr4 = cashflow.Cells.Find(c => c.Key == "I24").Value,
                freeCapExpYr5 = cashflow.Cells.Find(c => c.Key == "J24").Value,

                freeCapExpBorrowYr1 = cashflow.Cells.Find(c => c.Key == "F25").Value,
                freeCapExpBorrowYr2 = cashflow.Cells.Find(c => c.Key == "G25").Value,
                freeCapExpBorrowYr3 = cashflow.Cells.Find(c => c.Key == "H25").Value,
                freeCapExpBorrowYr4 = cashflow.Cells.Find(c => c.Key == "I25").Value,
                freeCapExpBorrowYr5 = cashflow.Cells.Find(c => c.Key == "J25").Value,

                freeCapExpPayYr1 = cashflow.Cells.Find(c => c.Key == "F26").Value,
                freeCapExpPayYr2 = cashflow.Cells.Find(c => c.Key == "G26").Value,
                freeCapExpPayYr3 = cashflow.Cells.Find(c => c.Key == "H26").Value,
                freeCapExpPayYr4 = cashflow.Cells.Find(c => c.Key == "I26").Value,
                freeCapExpPayYr5 = cashflow.Cells.Find(c => c.Key == "J26").Value,

                freeEarnOutPayPriceAdjYr1 = cashflow.Cells.Find(c => c.Key == "F27").Value,
                freeEarnOutPayPriceAdjYr2 = cashflow.Cells.Find(c => c.Key == "G27").Value,
                freeEarnOutPayPriceAdjYr3 = cashflow.Cells.Find(c => c.Key == "H27").Value,
                freeEarnOutPayPriceAdjYr4 = cashflow.Cells.Find(c => c.Key == "I27").Value,
                freeEarnOutPayPriceAdjYr5 = cashflow.Cells.Find(c => c.Key == "J27").Value,

                freeDisShareholdertaxyr1 = cashflow.Cells.Find(c => c.Key == "F28").Value,
                freeDisShareholdertaxyr2 = cashflow.Cells.Find(c => c.Key == "G28").Value,
                freeDisShareholdertaxyr3 = cashflow.Cells.Find(c => c.Key == "H28").Value,
                freeDisShareholdertaxyr4 = cashflow.Cells.Find(c => c.Key == "I28").Value,
                freeDisShareholdertaxyr5 = cashflow.Cells.Find(c => c.Key == "J28").Value,

                freeOpeCashFlowBusiYr1 = cashflow.Cells.Find(c => c.Key == "F29").Value,
                freeOpeCashFlowBusiYr2 = cashflow.Cells.Find(c => c.Key == "G29").Value,
                freeOpeCashFlowBusiYr3 = cashflow.Cells.Find(c => c.Key == "H29").Value,
                freeOpeCashFlowBusiYr4 = cashflow.Cells.Find(c => c.Key == "I29").Value,
                freeOpeCashFlowBusiYr5 = cashflow.Cells.Find(c => c.Key == "J29").Value,

                freeOpeCashFlowRealEstateYr1 = cashflow.Cells.Find(c => c.Key == "F30").Value,
                freeOpeCashFlowRealEstateYr2 = cashflow.Cells.Find(c => c.Key == "G30").Value,
                freeOpeCashFlowRealEstateYr3 = cashflow.Cells.Find(c => c.Key == "H30").Value,
                freeOpeCashFlowRealEstateYr4 = cashflow.Cells.Find(c => c.Key == "I30").Value,
                freeOpeCashFlowRealEstateYr5 = cashflow.Cells.Find(c => c.Key == "J30").Value,

                freeOperCashFlowTotalYr1 = cashflow.Cells.Find(c => c.Key == "F31").Value,
                freeOperCashFlowTotalYr2 = cashflow.Cells.Find(c => c.Key == "G31").Value,
                freeOperCashFlowTotalYr3 = cashflow.Cells.Find(c => c.Key == "H31").Value,
                freeOperCashFlowTotalYr4 = cashflow.Cells.Find(c => c.Key == "I31").Value,
                freeOperCashFlowTotalYr5 = cashflow.Cells.Find(c => c.Key == "J31").Value,

                freebeginCashBalYr1 = cashflow.Cells.Find(c => c.Key == "F34").Value,
                freebeginCashBalYr2 = cashflow.Cells.Find(c => c.Key == "G34").Value,
                freebeginCashBalYr3 = cashflow.Cells.Find(c => c.Key == "H34").Value,
                freebeginCashBalYr4 = cashflow.Cells.Find(c => c.Key == "I34").Value,
                freebeginCashBalYr5 = cashflow.Cells.Find(c => c.Key == "J34").Value,

                freeOperCashFlowYr1 = cashflow.Cells.Find(c => c.Key == "F35").Value,
                freeOperCashFlowYr2 = cashflow.Cells.Find(c => c.Key == "G35").Value,
                freeOperCashFlowYr3 = cashflow.Cells.Find(c => c.Key == "H35").Value,
                freeOperCashFlowYr4 = cashflow.Cells.Find(c => c.Key == "I35").Value,
                freeOperCashFlowYr5 = cashflow.Cells.Find(c => c.Key == "J35").Value,

                freeOpeCashReqYr1 = cashflow.Cells.Find(c => c.Key == "F36").Value,
                freeOpeCashReqYr2 = cashflow.Cells.Find(c => c.Key == "G36").Value,
                freeOpeCashReqYr3 = cashflow.Cells.Find(c => c.Key == "H36").Value,
                freeOpeCashReqYr4 = cashflow.Cells.Find(c => c.Key == "I36").Value,
                freeOpeCashReqYr5 = cashflow.Cells.Find(c => c.Key == "J36").Value,

                freeCashRevEndYr1 = cashflow.Cells.Find(c => c.Key == "F37").Value,
                freeCashRevEndYr2 = cashflow.Cells.Find(c => c.Key == "G37").Value,
                freeCashRevEndYr3 = cashflow.Cells.Find(c => c.Key == "H37").Value,
                freeCashRevEndYr4 = cashflow.Cells.Find(c => c.Key == "I37").Value,
                freeCashRevEndYr5 = cashflow.Cells.Find(c => c.Key == "J37").Value,

                freeb4BorrowingYr1 = cashflow.Cells.Find(c => c.Key == "F38").Value,
                freeb4BorrowingYr2 = cashflow.Cells.Find(c => c.Key == "G38").Value,
                freeb4BorrowingYr3 = cashflow.Cells.Find(c => c.Key == "H38").Value,
                freeb4BorrowingYr4 = cashflow.Cells.Find(c => c.Key == "I38").Value,
                freeb4BorrowingYr5 = cashflow.Cells.Find(c => c.Key == "J38").Value,

                freeAvailCreditLineyr1 = cashflow.Cells.Find(c => c.Key == "F40").Value,
                freeAvailCreditLineyr2 = cashflow.Cells.Find(c => c.Key == "G40").Value,
                freeAvailCreditLineyr3 = cashflow.Cells.Find(c => c.Key == "H40").Value,
                freeAvailCreditLineyr4 = cashflow.Cells.Find(c => c.Key == "I40").Value,
                freeAvailCreditLineyr5 = cashflow.Cells.Find(c => c.Key == "J40").Value,

                freeAddRevolverYr1 = cashflow.Cells.Find(c => c.Key == "F41").Value,
                freeAddRevolverYr2 = cashflow.Cells.Find(c => c.Key == "G41").Value,
                freeAddRevolverYr3 = cashflow.Cells.Find(c => c.Key == "H41").Value,
                freeAddRevolverYr4 = cashflow.Cells.Find(c => c.Key == "I41").Value,
                freeAddRevolverYr5 = cashflow.Cells.Find(c => c.Key == "J41").Value,

                freeBVXCashFlowYr1 = cashflow.Cells.Find(c => c.Key == "F43").Value,
                freeBVXCashFlowYr2 = cashflow.Cells.Find(c => c.Key == "G43").Value,
                freeBVXCashFlowYr3 = cashflow.Cells.Find(c => c.Key == "H43").Value,
                freeBVXCashFlowYr4 = cashflow.Cells.Find(c => c.Key == "I43").Value,
                freeBVXCashFlowYr5 = cashflow.Cells.Find(c => c.Key == "J43").Value,

                freeAddCapContributeYr1 = cashflow.Cells.Find(c => c.Key == "F46").Value,
                freeAddCapContributeYr2 = cashflow.Cells.Find(c => c.Key == "G46").Value,
                freeAddCapContributeYr3 = cashflow.Cells.Find(c => c.Key == "H46").Value,
                freeAddCapContributeYr4 = cashflow.Cells.Find(c => c.Key == "I46").Value,
                freeAddCapContributeYr5 = cashflow.Cells.Find(c => c.Key == "J46").Value,

                freeDividendDistrRegYr1 = cashflow.Cells.Find(c => c.Key == "F47").Value,
                freeDividendDistrRegYr2 = cashflow.Cells.Find(c => c.Key == "G47").Value,
                freeDividendDistrRegYr3 = cashflow.Cells.Find(c => c.Key == "H47").Value,
                freeDividendDistrRegYr4 = cashflow.Cells.Find(c => c.Key == "I47").Value,
                freeDividendDistrRegYr5 = cashflow.Cells.Find(c => c.Key == "J47").Value,

                freeAddOverAdvLoanPayYr1 = cashflow.Cells.Find(c => c.Key == "F48").Value,
                freeAddOverAdvLoanPayYr2 = cashflow.Cells.Find(c => c.Key == "G48").Value,
                freeAddOverAdvLoanPayYr3 = cashflow.Cells.Find(c => c.Key == "H48").Value,
                freeAddOverAdvLoanPayYr4 = cashflow.Cells.Find(c => c.Key == "I48").Value,
                freeAddOverAdvLoanPayYr5 = cashflow.Cells.Find(c => c.Key == "J48").Value,

                freeAddRevolverPayYr1 = cashflow.Cells.Find(c => c.Key == "F49").Value,
                freeAddRevolverPayYr2 = cashflow.Cells.Find(c => c.Key == "G49").Value,
                freeAddRevolverPayYr3 = cashflow.Cells.Find(c => c.Key == "H49").Value,
                freeAddRevolverPayYr4 = cashflow.Cells.Find(c => c.Key == "I49").Value,
                freeAddRevolverPayYr5 = cashflow.Cells.Find(c => c.Key == "J49").Value,

                freeAddTremloanPayYr1 = cashflow.Cells.Find(c => c.Key == "F50").Value,
                freeAddTremloanPayYr2 = cashflow.Cells.Find(c => c.Key == "G50").Value,
                freeAddTremloanPayYr3 = cashflow.Cells.Find(c => c.Key == "H50").Value,
                freeAddTremloanPayYr4 = cashflow.Cells.Find(c => c.Key == "I50").Value,
                freeAddTremloanPayYr5 = cashflow.Cells.Find(c => c.Key == "J50").Value,

                freeAddNewCapExPayYr1 = cashflow.Cells.Find(c => c.Key == "F51").Value,
                freeAddNewCapExPayYr2 = cashflow.Cells.Find(c => c.Key == "G51").Value,
                freeAddNewCapExPayYr3 = cashflow.Cells.Find(c => c.Key == "H51").Value,
                freeAddNewCapExPayYr4 = cashflow.Cells.Find(c => c.Key == "I51").Value,
                freeAddNewCapExPayYr5 = cashflow.Cells.Find(c => c.Key == "J51").Value,

                freeAddgapNotePayYr1 = cashflow.Cells.Find(c => c.Key == "F52").Value,
                freeAddgapNotePayYr2 = cashflow.Cells.Find(c => c.Key == "G52").Value,
                freeAddgapNotePayYr3 = cashflow.Cells.Find(c => c.Key == "H52").Value,
                freeAddgapNotePayYr4 = cashflow.Cells.Find(c => c.Key == "I52").Value,
                freeAddgapNotePayYr5 = cashflow.Cells.Find(c => c.Key == "J52").Value,

                freeDiviDistrAddYr1 = cashflow.Cells.Find(c => c.Key == "F53").Value,
                freeDiviDistrAddYr2 = cashflow.Cells.Find(c => c.Key == "G53").Value,
                freeDiviDistrAddYr3 = cashflow.Cells.Find(c => c.Key == "H53").Value,
                freeDiviDistrAddYr4 = cashflow.Cells.Find(c => c.Key == "I53").Value,
                freeDiviDistrAddYr5 = cashflow.Cells.Find(c => c.Key == "J53").Value,

                freeCapitalFusionYr1 = cashflow.Cells.Find(c => c.Key == "F54").Value,
                freeCapitalFusionYr2 = cashflow.Cells.Find(c => c.Key == "G54").Value,
                freeCapitalFusionYr3 = cashflow.Cells.Find(c => c.Key == "H54").Value,
                freeCapitalFusionYr4 = cashflow.Cells.Find(c => c.Key == "I54").Value,
                freeCapitalFusionYr5 = cashflow.Cells.Find(c => c.Key == "J54").Value,

                freeChangeinCashYr1 = cashflow.Cells.Find(c => c.Key == "F56").Value,
                freeChangeinCashYr2 = cashflow.Cells.Find(c => c.Key == "G56").Value,
                freeChangeinCashYr3 = cashflow.Cells.Find(c => c.Key == "H56").Value,
                freeChangeinCashYr4 = cashflow.Cells.Find(c => c.Key == "I56").Value,
                freeChangeinCashYr5 = cashflow.Cells.Find(c => c.Key == "J56").Value,

                // Cash Flow Statement

                stmtNetIncomeYr1 = cashstate.Cells.Find(c => c.Key == "F9").Value,
                stmtNetIncomeYr2 = cashstate.Cells.Find(c => c.Key == "G9").Value,
                stmtNetIncomeYr3 = cashstate.Cells.Find(c => c.Key == "H9").Value,
                stmtNetIncomeYr4 = cashstate.Cells.Find(c => c.Key == "I9").Value,
                stmtNetIncomeYr5 = cashstate.Cells.Find(c => c.Key == "J9").Value,

                stmtDepriciationYr1 = cashstate.Cells.Find(c => c.Key == "F10").Value,
                stmtDepriciationYr2 = cashstate.Cells.Find(c => c.Key == "G10").Value,
                stmtDepriciationYr3 = cashstate.Cells.Find(c => c.Key == "H10").Value,
                stmtDepriciationYr4 = cashstate.Cells.Find(c => c.Key == "I10").Value,
                stmtDepriciationYr5 = cashstate.Cells.Find(c => c.Key == "J10").Value,

                stmtnonCompAmortyr1 = cashstate.Cells.Find(c => c.Key == "F11").Value,
                stmtnonCompAmortyr2 = cashstate.Cells.Find(c => c.Key == "G11").Value,
                stmtnonCompAmortyr3 = cashstate.Cells.Find(c => c.Key == "H11").Value,
                stmtnonCompAmortyr4 = cashstate.Cells.Find(c => c.Key == "I11").Value,
                stmtnonCompAmortyr5 = cashstate.Cells.Find(c => c.Key == "J11").Value,

                stmtPerGoodwillAmortYr1 = cashstate.Cells.Find(c => c.Key == "F12").Value,
                stmtPerGoodwillAmortYr2 = cashstate.Cells.Find(c => c.Key == "G12").Value,
                stmtPerGoodwillAmortYr3 = cashstate.Cells.Find(c => c.Key == "H12").Value,
                stmtPerGoodwillAmortYr4 = cashstate.Cells.Find(c => c.Key == "I12").Value,
                stmtPerGoodwillAmortYr5 = cashstate.Cells.Find(c => c.Key == "J12").Value,

                stmtPreConAmortYr1 = cashstate.Cells.Find(c => c.Key == "F13").Value,
                stmtPreConAmortYr2 = cashstate.Cells.Find(c => c.Key == "G13").Value,
                stmtPreConAmortYr3 = cashstate.Cells.Find(c => c.Key == "H13").Value,
                stmtPreConAmortYr4 = cashstate.Cells.Find(c => c.Key == "I13").Value,
                stmtPreConAmortYr5 = cashstate.Cells.Find(c => c.Key == "J13").Value,

                stmtAcqCostAmortYr1 = cashstate.Cells.Find(c => c.Key == "F14").Value,
                stmtAcqCostAmortYr2 = cashstate.Cells.Find(c => c.Key == "G14").Value,
                stmtAcqCostAmortYr3 = cashstate.Cells.Find(c => c.Key == "H14").Value,
                stmtAcqCostAmortYr4 = cashstate.Cells.Find(c => c.Key == "I14").Value,
                stmtAcqCostAmortYr5 = cashstate.Cells.Find(c => c.Key == "J14").Value,

                stmtGoodwillAmortYr1 = cashstate.Cells.Find(c => c.Key == "F15").Value,
                stmtGoodwillAmortYr2 = cashstate.Cells.Find(c => c.Key == "G15").Value,
                stmtGoodwillAmortYr3 = cashstate.Cells.Find(c => c.Key == "H15").Value,
                stmtGoodwillAmortYr4 = cashstate.Cells.Find(c => c.Key == "I15").Value,
                stmtGoodwillAmortYr5 = cashstate.Cells.Find(c => c.Key == "J15").Value,

                stmtWChangeYr1 = cashstate.Cells.Find(c => c.Key == "F16").Value,
                stmtWChangeYr2 = cashstate.Cells.Find(c => c.Key == "G16").Value,
                stmtWChangeYr3 = cashstate.Cells.Find(c => c.Key == "H16").Value,
                stmtWChangeYr4 = cashstate.Cells.Find(c => c.Key == "I16").Value,
                stmtWChangeYr5 = cashstate.Cells.Find(c => c.Key == "J16").Value,

                stmtOperActivityYr1 = cashstate.Cells.Find(c => c.Key == "F17").Value,
                stmtOperActivityYr2 = cashstate.Cells.Find(c => c.Key == "G17").Value,
                stmtOperActivityYr3 = cashstate.Cells.Find(c => c.Key == "H17").Value,
                stmtOperActivityYr4 = cashstate.Cells.Find(c => c.Key == "I17").Value,
                stmtOperActivityYr5 = cashstate.Cells.Find(c => c.Key == "J17").Value,

                stmtCapitalExpYr1 = cashstate.Cells.Find(c => c.Key == "F20").Value,
                stmtCapitalExpYr2 = cashstate.Cells.Find(c => c.Key == "G20").Value,
                stmtCapitalExpYr3 = cashstate.Cells.Find(c => c.Key == "H20").Value,
                stmtCapitalExpYr4 = cashstate.Cells.Find(c => c.Key == "I20").Value,
                stmtCapitalExpYr5 = cashstate.Cells.Find(c => c.Key == "J20").Value,

                stmtPriceAdjYr1 = cashstate.Cells.Find(c => c.Key == "F21").Value,
                stmtPriceAdjYr2 = cashstate.Cells.Find(c => c.Key == "G21").Value,
                stmtPriceAdjYr3 = cashstate.Cells.Find(c => c.Key == "H21").Value,
                stmtPriceAdjYr4 = cashstate.Cells.Find(c => c.Key == "I21").Value,
                stmtPriceAdjYr5 = cashstate.Cells.Find(c => c.Key == "J21").Value,

                stmtRemNonCompPayYr1 = cashstate.Cells.Find(c => c.Key == "F22").Value,
                stmtRemNonCompPayYr2 = cashstate.Cells.Find(c => c.Key == "G22").Value,
                stmtRemNonCompPayYr3 = cashstate.Cells.Find(c => c.Key == "H22").Value,
                stmtRemNonCompPayYr4 = cashstate.Cells.Find(c => c.Key == "I22").Value,
                stmtRemNonCompPayYr5 = cashstate.Cells.Find(c => c.Key == "J22").Value,

                stmtRemPerGWPayYr1 = cashstate.Cells.Find(c => c.Key == "F23").Value,
                stmtRemPerGWPayYr2 = cashstate.Cells.Find(c => c.Key == "G23").Value,
                stmtRemPerGWPayYr3 = cashstate.Cells.Find(c => c.Key == "H23").Value,
                stmtRemPerGWPayYr4 = cashstate.Cells.Find(c => c.Key == "I23").Value,
                stmtRemPerGWPayYr5 = cashstate.Cells.Find(c => c.Key == "J23").Value,

                stmtInvActYr1 = cashstate.Cells.Find(c => c.Key == "F24").Value,
                stmtInvActYr2 = cashstate.Cells.Find(c => c.Key == "G24").Value,
                stmtInvActYr3 = cashstate.Cells.Find(c => c.Key == "H24").Value,
                stmtInvActYr4 = cashstate.Cells.Find(c => c.Key == "I24").Value,
                stmtInvActYr5 = cashstate.Cells.Find(c => c.Key == "J24").Value,

                stmtAddRevolverYr1 = cashstate.Cells.Find(c => c.Key == "F27").Value,
                stmtAddRevolverYr2 = cashstate.Cells.Find(c => c.Key == "G27").Value,
                stmtAddRevolverYr3 = cashstate.Cells.Find(c => c.Key == "H27").Value,
                stmtAddRevolverYr4 = cashstate.Cells.Find(c => c.Key == "I27").Value,
                stmtAddRevolverYr5 = cashstate.Cells.Find(c => c.Key == "J27").Value,

                stmtRevPayWCYr1 = cashstate.Cells.Find(c => c.Key == "F28").Value,
                stmtRevPayWCYr2 = cashstate.Cells.Find(c => c.Key == "G28").Value,
                stmtRevPayWCYr3 = cashstate.Cells.Find(c => c.Key == "H28").Value,
                stmtRevPayWCYr4 = cashstate.Cells.Find(c => c.Key == "I28").Value,
                stmtRevPayWCYr5 = cashstate.Cells.Find(c => c.Key == "J28").Value,

                stmtAddRevPayYr1 = cashstate.Cells.Find(c => c.Key == "F29").Value,
                stmtAddRevPayYr2 = cashstate.Cells.Find(c => c.Key == "G29").Value,
                stmtAddRevPayYr3 = cashstate.Cells.Find(c => c.Key == "H29").Value,
                stmtAddRevPayYr4 = cashstate.Cells.Find(c => c.Key == "I29").Value,
                stmtAddRevPayYr5 = cashstate.Cells.Find(c => c.Key == "J29").Value,

                stmtTermLoanPayYr1 = cashstate.Cells.Find(c => c.Key == "F30").Value,
                stmtTermLoanPayYr2 = cashstate.Cells.Find(c => c.Key == "G30").Value,
                stmtTermLoanPayYr3 = cashstate.Cells.Find(c => c.Key == "H30").Value,
                stmtTermLoanPayYr4 = cashstate.Cells.Find(c => c.Key == "I30").Value,
                stmtTermLoanPayYr5 = cashstate.Cells.Find(c => c.Key == "J30").Value,

                stmtAddTermLoanpayYr1 = cashstate.Cells.Find(c => c.Key == "F31").Value,
                stmtAddTermLoanpayYr2 = cashstate.Cells.Find(c => c.Key == "G31").Value,
                stmtAddTermLoanpayYr3 = cashstate.Cells.Find(c => c.Key == "H31").Value,
                stmtAddTermLoanpayYr4 = cashstate.Cells.Find(c => c.Key == "I31").Value,
                stmtAddTermLoanpayYr5 = cashstate.Cells.Find(c => c.Key == "J31").Value,

                stmtOverAdvLoanPayYr1 = cashstate.Cells.Find(c => c.Key == "F32").Value,
                stmtOverAdvLoanPayYr2 = cashstate.Cells.Find(c => c.Key == "G32").Value,
                stmtOverAdvLoanPayYr3 = cashstate.Cells.Find(c => c.Key == "H32").Value,
                stmtOverAdvLoanPayYr4 = cashstate.Cells.Find(c => c.Key == "I32").Value,
                stmtOverAdvLoanPayYr5 = cashstate.Cells.Find(c => c.Key == "J32").Value,

                stmtAddOverAdvLoanPayYr1 = cashstate.Cells.Find(c => c.Key == "F33").Value,
                stmtAddOverAdvLoanPayYr2 = cashstate.Cells.Find(c => c.Key == "G33").Value,
                stmtAddOverAdvLoanPayYr3 = cashstate.Cells.Find(c => c.Key == "H33").Value,
                stmtAddOverAdvLoanPayYr4 = cashstate.Cells.Find(c => c.Key == "I33").Value,
                stmtAddOverAdvLoanPayYr5 = cashstate.Cells.Find(c => c.Key == "J33").Value,

                stmtMezzFinanPayYr1 = cashstate.Cells.Find(c => c.Key == "F34").Value,
                stmtMezzFinanPayYr2 = cashstate.Cells.Find(c => c.Key == "G34").Value,
                stmtMezzFinanPayYr3 = cashstate.Cells.Find(c => c.Key == "H34").Value,
                stmtMezzFinanPayYr4 = cashstate.Cells.Find(c => c.Key == "I34").Value,
                stmtMezzFinanPayYr5 = cashstate.Cells.Find(c => c.Key == "J34").Value,

                stmtGapNotePayYr1 = cashstate.Cells.Find(c => c.Key == "F35").Value,
                stmtGapNotePayYr2 = cashstate.Cells.Find(c => c.Key == "G35").Value,
                stmtGapNotePayYr3 = cashstate.Cells.Find(c => c.Key == "H35").Value,
                stmtGapNotePayYr4 = cashstate.Cells.Find(c => c.Key == "I35").Value,
                stmtGapNotePayYr5 = cashstate.Cells.Find(c => c.Key == "J35").Value,

                stmtGapBalloonPayYr1 = cashstate.Cells.Find(c => c.Key == "F36").Value,
                stmtGapBalloonPayYr2 = cashstate.Cells.Find(c => c.Key == "G36").Value,
                stmtGapBalloonPayYr3 = cashstate.Cells.Find(c => c.Key == "H36").Value,
                stmtGapBalloonPayYr4 = cashstate.Cells.Find(c => c.Key == "I36").Value,
                stmtGapBalloonPayYr5 = cashstate.Cells.Find(c => c.Key == "J36").Value,

                stmtAddGapNotePayYr1 = cashstate.Cells.Find(c => c.Key == "F37").Value,
                stmtAddGapNotePayYr2 = cashstate.Cells.Find(c => c.Key == "G37").Value,
                stmtAddGapNotePayYr3 = cashstate.Cells.Find(c => c.Key == "H37").Value,
                stmtAddGapNotePayYr4 = cashstate.Cells.Find(c => c.Key == "I37").Value,
                stmtAddGapNotePayYr5 = cashstate.Cells.Find(c => c.Key == "J37").Value,

                stmtCapExpBorrowYr1 = cashstate.Cells.Find(c => c.Key == "F38").Value,
                stmtCapExpBorrowYr2 = cashstate.Cells.Find(c => c.Key == "G38").Value,
                stmtCapExpBorrowYr3 = cashstate.Cells.Find(c => c.Key == "H38").Value,
                stmtCapExpBorrowYr4 = cashstate.Cells.Find(c => c.Key == "I38").Value,
                stmtCapExpBorrowYr5 = cashstate.Cells.Find(c => c.Key == "J38").Value,

                stmtCapExpPayYr1 = cashstate.Cells.Find(c => c.Key == "F39").Value,
                stmtCapExpPayYr2 = cashstate.Cells.Find(c => c.Key == "G39").Value,
                stmtCapExpPayYr3 = cashstate.Cells.Find(c => c.Key == "H39").Value,
                stmtCapExpPayYr4 = cashstate.Cells.Find(c => c.Key == "I39").Value,
                stmtCapExpPayYr5 = cashstate.Cells.Find(c => c.Key == "J39").Value,

                stmtAddNewCapExPayYr1 = cashstate.Cells.Find(c => c.Key == "F40").Value,
                stmtAddNewCapExPayYr2 = cashstate.Cells.Find(c => c.Key == "G40").Value,
                stmtAddNewCapExPayYr3 = cashstate.Cells.Find(c => c.Key == "H40").Value,
                stmtAddNewCapExPayYr4 = cashstate.Cells.Find(c => c.Key == "I40").Value,
                stmtAddNewCapExPayYr5 = cashstate.Cells.Find(c => c.Key == "J40").Value,

                stmtAddCapContributeYr1 = cashstate.Cells.Find(c => c.Key == "F41").Value,
                stmtAddCapContributeYr2 = cashstate.Cells.Find(c => c.Key == "G41").Value,
                stmtAddCapContributeYr3 = cashstate.Cells.Find(c => c.Key == "H41").Value,
                stmtAddCapContributeYr4 = cashstate.Cells.Find(c => c.Key == "I41").Value,
                stmtAddCapContributeYr5 = cashstate.Cells.Find(c => c.Key == "J41").Value,

                stmtDividendDisRegYr1 = cashstate.Cells.Find(c => c.Key == "F42").Value,
                stmtDividendDisRegYr2 = cashstate.Cells.Find(c => c.Key == "G42").Value,
                stmtDividendDisRegYr3 = cashstate.Cells.Find(c => c.Key == "H42").Value,
                stmtDividendDisRegYr4 = cashstate.Cells.Find(c => c.Key == "I42").Value,
                stmtDividendDisRegYr5 = cashstate.Cells.Find(c => c.Key == "J42").Value,

                stmtDividendDisAddYr1 = cashstate.Cells.Find(c => c.Key == "F43").Value,
                stmtDividendDisAddYr2 = cashstate.Cells.Find(c => c.Key == "G43").Value,
                stmtDividendDisAddYr3 = cashstate.Cells.Find(c => c.Key == "H43").Value,
                stmtDividendDisAddYr4 = cashstate.Cells.Find(c => c.Key == "I43").Value,
                stmtDividendDisAddYr5 = cashstate.Cells.Find(c => c.Key == "J43").Value,

                stmtDisrShareTaxYr1 = cashstate.Cells.Find(c => c.Key == "F44").Value,
                stmtDisrShareTaxYr2 = cashstate.Cells.Find(c => c.Key == "G44").Value,
                stmtDisrShareTaxYr3 = cashstate.Cells.Find(c => c.Key == "H44").Value,
                stmtDisrShareTaxYr4 = cashstate.Cells.Find(c => c.Key == "I44").Value,
                stmtDisrShareTaxYr5 = cashstate.Cells.Find(c => c.Key == "J44").Value,

                stmtFinanActivitiesYr1 = cashstate.Cells.Find(c => c.Key == "F45").Value,
                stmtFinanActivitiesYr2 = cashstate.Cells.Find(c => c.Key == "G45").Value,
                stmtFinanActivitiesYr3 = cashstate.Cells.Find(c => c.Key == "H45").Value,
                stmtFinanActivitiesYr4 = cashstate.Cells.Find(c => c.Key == "I45").Value,
                stmtFinanActivitiesYr5 = cashstate.Cells.Find(c => c.Key == "J45").Value,

                stmtCashBusinessYr1 = cashstate.Cells.Find(c => c.Key == "F47").Value,
                stmtCashBusinessYr2 = cashstate.Cells.Find(c => c.Key == "G47").Value,
                stmtCashBusinessYr3 = cashstate.Cells.Find(c => c.Key == "H47").Value,
                stmtCashBusinessYr4 = cashstate.Cells.Find(c => c.Key == "I47").Value,
                stmtCashBusinessYr5 = cashstate.Cells.Find(c => c.Key == "J47").Value,

                stmtCashEstateYr1 = cashstate.Cells.Find(c => c.Key == "F48").Value,
                stmtCashEstateYr2 = cashstate.Cells.Find(c => c.Key == "G48").Value,
                stmtCashEstateYr3 = cashstate.Cells.Find(c => c.Key == "H48").Value,
                stmtCashEstateYr4 = cashstate.Cells.Find(c => c.Key == "I48").Value,
                stmtCashEstateYr5 = cashstate.Cells.Find(c => c.Key == "J48").Value,

                stmtTotalinCashYr1 = cashstate.Cells.Find(c => c.Key == "F49").Value,
                stmtTotalinCashYr2 = cashstate.Cells.Find(c => c.Key == "G49").Value,
                stmtTotalinCashYr3 = cashstate.Cells.Find(c => c.Key == "H49").Value,
                stmtTotalinCashYr4 = cashstate.Cells.Find(c => c.Key == "I49").Value,
                stmtTotalinCashYr5 = cashstate.Cells.Find(c => c.Key == "J49").Value,

                stmtBeginCashBalYr1 = cashstate.Cells.Find(c => c.Key == "F51").Value,
                stmtBeginCashBalYr2 = cashstate.Cells.Find(c => c.Key == "G51").Value,
                stmtBeginCashBalYr3 = cashstate.Cells.Find(c => c.Key == "H51").Value,
                stmtBeginCashBalYr4 = cashstate.Cells.Find(c => c.Key == "I51").Value,
                stmtBeginCashBalYr5 = cashstate.Cells.Find(c => c.Key == "J51").Value,

                stmtEndCashBalYr1 = cashstate.Cells.Find(c => c.Key == "F53").Value,
                stmtEndCashBalYr2 = cashstate.Cells.Find(c => c.Key == "G53").Value,
                stmtEndCashBalYr3 = cashstate.Cells.Find(c => c.Key == "H53").Value,
                stmtEndCashBalYr4 = cashstate.Cells.Find(c => c.Key == "I53").Value,
                stmtEndCashBalYr5 = cashstate.Cells.Find(c => c.Key == "J53").Value,

                // ROI Page

                roiOrigEquInvYr0 = roi.Cells.Find(c => c.Key == "E8").Value,

                roiExitMultipleYr5 = roi.Cells.Find(c => c.Key == "J9").Value,

                roiPlusCashRetainYr5 = roi.Cells.Find(c => c.Key == "J10").Value,

                roiClosingCostExitYr5 = roi.Cells.Find(c => c.Key == "J11").Value,

                roiStateYr5 = roi.Cells.Find(c => c.Key == "J12").Value,

                roiLessNonOperLiabYr5 = roi.Cells.Find(c => c.Key == "J13").Value,

                roiTerminalValYr5 = roi.Cells.Find(c => c.Key == "J14").Value,

                roiShareholderYr1 = roi.Cells.Find(c => c.Key == "F16").Value,
                roiShareholderYr2 = roi.Cells.Find(c => c.Key == "G16").Value,
                roiShareholderYr3 = roi.Cells.Find(c => c.Key == "H16").Value,
                roiShareholderYr4 = roi.Cells.Find(c => c.Key == "I16").Value,
                roiShareholderYr5 = roi.Cells.Find(c => c.Key == "J16").Value,

                roiIRSYr1 = roi.Cells.Find(c => c.Key == "F17").Value,
                roiIRSYr2 = roi.Cells.Find(c => c.Key == "G17").Value,
                roiIRSYr3 = roi.Cells.Find(c => c.Key == "H17").Value,
                roiIRSYr4 = roi.Cells.Find(c => c.Key == "I17").Value,
                roiIRSYr5 = roi.Cells.Find(c => c.Key == "J17").Value,

                roiAddCapContribution1 = roi.Cells.Find(c => c.Key == "F18").Value,
                roiAddCapContribution2 = roi.Cells.Find(c => c.Key == "G18").Value,
                roiAddCapContribution3 = roi.Cells.Find(c => c.Key == "H18").Value,
                roiAddCapContribution4 = roi.Cells.Find(c => c.Key == "I18").Value,
                roiAddCapContribution5 = roi.Cells.Find(c => c.Key == "J18").Value,

                roiDividendDist1 = roi.Cells.Find(c => c.Key == "F19").Value,
                roiDividendDist2 = roi.Cells.Find(c => c.Key == "G19").Value,
                roiDividendDist3 = roi.Cells.Find(c => c.Key == "H19").Value,
                roiDividendDist4 = roi.Cells.Find(c => c.Key == "I19").Value,
                roiDividendDist5 = roi.Cells.Find(c => c.Key == "J19").Value,

                roiSdivPreTaxYr1 = roi.Cells.Find(c => c.Key == "F20").Value,
                roiSdivPreTaxYr2 = roi.Cells.Find(c => c.Key == "G20").Value,
                roiSdivPreTaxYr3 = roi.Cells.Find(c => c.Key == "H20").Value,
                roiSdivPreTaxYr4 = roi.Cells.Find(c => c.Key == "I20").Value,
                roiSdivPreTaxYr5 = roi.Cells.Find(c => c.Key == "J20").Value,

                roiUndistributeYr1 = roi.Cells.Find(c => c.Key == "F21").Value,
                roiUndistributeYr2 = roi.Cells.Find(c => c.Key == "G21").Value,
                roiUndistributeYr3 = roi.Cells.Find(c => c.Key == "H21").Value,
                roiUndistributeYr4 = roi.Cells.Find(c => c.Key == "I21").Value,
                roiUndistributeYr5 = roi.Cells.Find(c => c.Key == "J21").Value,

                roiblankYr1 = roi.Cells.Find(c => c.Key == "F22").Value,
                roiblankYr2 = roi.Cells.Find(c => c.Key == "G22").Value,
                roiblankYr3 = roi.Cells.Find(c => c.Key == "H22").Value,
                roiblankYr4 = roi.Cells.Find(c => c.Key == "I22").Value,
                roiblankYr5 = roi.Cells.Find(c => c.Key == "J22").Value,

                roiByerCashFlowyr1 = roi.Cells.Find(c => c.Key == "F24").Value,
                roiByerCashFlowyr2 = roi.Cells.Find(c => c.Key == "G24").Value,
                roiByerCashFlowyr3 = roi.Cells.Find(c => c.Key == "H24").Value,
                roiByerCashFlowyr4 = roi.Cells.Find(c => c.Key == "I24").Value,
                roiByerCashFlowyr5 = roi.Cells.Find(c => c.Key == "J24").Value,

                roiMezzCashFlowYr0 = roi.Cells.Find(c => c.Key == "E25").Value,
                roiMezzCashFlowYr1 = roi.Cells.Find(c => c.Key == "F25").Value,
                roiMezzCashFlowYr2 = roi.Cells.Find(c => c.Key == "G25").Value,
                roiMezzCashFlowYr3 = roi.Cells.Find(c => c.Key == "H25").Value,
                roiMezzCashFlowYr4 = roi.Cells.Find(c => c.Key == "I25").Value,
                roiMezzCashFlowYr5 = roi.Cells.Find(c => c.Key == "J25").Value,

                roiByerPreTaxyr0 = roi.Cells.Find(c => c.Key == "E26").Value,
                roiByerPreTaxyr1 = roi.Cells.Find(c => c.Key == "F26").Value,
                roiByerPreTaxyr2 = roi.Cells.Find(c => c.Key == "G26").Value,
                roiByerPreTaxyr3 = roi.Cells.Find(c => c.Key == "H26").Value,
                roiByerPreTaxyr4 = roi.Cells.Find(c => c.Key == "I26").Value,
                roiByerPreTaxyr5 = roi.Cells.Find(c => c.Key == "J26").Value,

                roiByerPreTaxROE = roi.Cells.Find(c => c.Key == "F28").Value,

                roiMezzFinancing0 = roi.Cells.Find(c => c.Key == "E34").Value,

                roiIntExpMezzFinanYr0 = roi.Cells.Find(c => c.Key == "E35").Value,
                roiIntExpMezzFinanYr1 = roi.Cells.Find(c => c.Key == "F35").Value,
                roiIntExpMezzFinanYr2 = roi.Cells.Find(c => c.Key == "G35").Value,
                roiIntExpMezzFinanYr3 = roi.Cells.Find(c => c.Key == "H35").Value,
                roiIntExpMezzFinanYr4 = roi.Cells.Find(c => c.Key == "I35").Value,
                roiIntExpMezzFinanYr5 = roi.Cells.Find(c => c.Key == "J35").Value,

                roiMezzPrincipalAmortYr0 = roi.Cells.Find(c => c.Key == "E36").Value,
                roiMezzPrincipalAmortYr1 = roi.Cells.Find(c => c.Key == "F36").Value,
                roiMezzPrincipalAmortYr2 = roi.Cells.Find(c => c.Key == "G36").Value,
                roiMezzPrincipalAmortYr3 = roi.Cells.Find(c => c.Key == "H36").Value,
                roiMezzPrincipalAmortYr4 = roi.Cells.Find(c => c.Key == "I36").Value,
                roiMezzPrincipalAmortYr5 = roi.Cells.Find(c => c.Key == "J36").Value,

                roiMezzRemPrincePayYr0 = roi.Cells.Find(c => c.Key == "E37").Value,
                roiMezzRemPrincePayYr1 = roi.Cells.Find(c => c.Key == "F37").Value,
                roiMezzRemPrincePayYr2 = roi.Cells.Find(c => c.Key == "G37").Value,
                roiMezzRemPrincePayYr3 = roi.Cells.Find(c => c.Key == "H37").Value,
                roiMezzRemPrincePayYr4 = roi.Cells.Find(c => c.Key == "I37").Value,
                roiMezzRemPrincePayYr5 = roi.Cells.Find(c => c.Key == "J37").Value,

                roiMezzSharePreTaxYr0 = roi.Cells.Find(c => c.Key == "E38").Value,
                roiMezzSharePreTaxYr1 = roi.Cells.Find(c => c.Key == "F38").Value,
                roiMezzSharePreTaxYr2 = roi.Cells.Find(c => c.Key == "G38").Value,
                roiMezzSharePreTaxYr3 = roi.Cells.Find(c => c.Key == "H38").Value,
                roiMezzSharePreTaxYr4 = roi.Cells.Find(c => c.Key == "I38").Value,
                roiMezzSharePreTaxYr5 = roi.Cells.Find(c => c.Key == "J38").Value,

                roiMezzaPreTaxCasshFlowYr0 = roi.Cells.Find(c => c.Key == "E39").Value,
                roiMezzaPreTaxCasshFlowYr1 = roi.Cells.Find(c => c.Key == "F39").Value,
                roiMezzaPreTaxCasshFlowYr2 = roi.Cells.Find(c => c.Key == "G39").Value,
                roiMezzaPreTaxCasshFlowYr3 = roi.Cells.Find(c => c.Key == "H39").Value,
                roiMezzaPreTaxCasshFlowYr4 = roi.Cells.Find(c => c.Key == "I39").Value,
                roiMezzaPreTaxCasshFlowYr5 = roi.Cells.Find(c => c.Key == "J39").Value,

                roiActualMezzPreTaxROIYr0 = roi.Cells.Find(c => c.Key == "E41").Value,

                roiExpectMezzYr0 = roi.Cells.Find(c => c.Key == "E42").Value,


            };
            CultureInfo cultures = new CultureInfo("en-US");

            respone.ebitdaMultiple = Math.Round(Convert.ToDecimal(respone.ebitdaMultiple, cultures), 2).ToString();
            respone.ebitMultiple = Math.Round(Convert.ToDecimal(respone.ebitMultiple, cultures), 2).ToString();
            respone.ebitda = Math.Round(Convert.ToDecimal(respone.ebitda, cultures), 2).ToString("#,####");
            respone.bvAsset = Math.Round(Convert.ToDecimal(respone.bvAsset, cultures), 0).ToString("#,####");
            respone.fmvAsset = Math.Round(Convert.ToDecimal(respone.fmvAsset, cultures), 0).ToString("#,####");
            respone.goodwill = Math.Round(Convert.ToDecimal(respone.goodwill, cultures), 0).ToString("#,####");
            respone.ev = Math.Round(Convert.ToDecimal(respone.ev, cultures), 0).ToString();
            respone.cashclosing = Math.Round(Convert.ToDecimal(respone.cashclosing, cultures), 2).ToString();
            respone.cashclosingper = Math.Round(Convert.ToDecimal(respone.cashclosingper, cultures), 2).ToString();
           // respone.sellernote = Math.Round(Convert.ToDecimal(respone.sellernote, cultures),2).ToString();
           // respone.sellernoteper = Math.Round(Convert.ToDecimal(respone.sellernoteper, cultures), 2).ToString();
            respone.evalue = Math.Round(Convert.ToDecimal(respone.evalue, cultures), 0).ToString();
            respone.buyerequ = Math.Round(Convert.ToDecimal(respone.buyerequ, cultures), 0).ToString();
            respone.buyerequper = Math.Round(Convert.ToDecimal(respone.buyerequper, cultures), 2).ToString();
            respone.buyerroe = Math.Round(Convert.ToDecimal(respone.buyerroe, cultures), 2).ToString();
            respone.buyerequity = Math.Round(Convert.ToDecimal(respone.buyerequity, cultures), 0).ToString();
            respone.revolverterm = Math.Round(Convert.ToDecimal(respone.revolverterm, cultures), 0).ToString();
            respone.totalcapraised = Math.Round(Convert.ToDecimal(respone.totalcapraised, cultures), 0).ToString();
            respone.lessacquisition = Math.Round(Convert.ToDecimal(respone.lessacquisition, cultures), 0).ToString();
            respone.cashtosellerclosing = Math.Round(Convert.ToDecimal(respone.cashtosellerclosing, cultures), 0).ToString();

            respone.sales0 = Math.Round(Convert.ToDecimal(respone.sales0, cultures), 0).ToString();
            respone.sales1 = Math.Round(Convert.ToDecimal(respone.sales1, cultures), 0).ToString();
            respone.sales2 = Math.Round(Convert.ToDecimal(respone.sales2, cultures), 0).ToString();
            respone.sales3 = Math.Round(Convert.ToDecimal(respone.sales3, cultures), 0).ToString();
            respone.sales4 = Math.Round(Convert.ToDecimal(respone.sales4, cultures), 0).ToString();
            respone.sales5 = Math.Round(Convert.ToDecimal(respone.sales5, cultures), 0).ToString();

            respone.growth1 = Math.Round(Convert.ToDecimal(respone.growth1, cultures), 2).ToString();
            respone.growth2 = Math.Round(Convert.ToDecimal(respone.growth2, cultures), 2).ToString();
            /*respone.growth3 = Math.Round(Convert.ToDecimal(respone.growth3, cultures), 2).ToString();
            respone.growth4 = Math.Round(Convert.ToDecimal(respone.growth4, cultures), 2).ToString();
            respone.growth5 = Math.Round(Convert.ToDecimal(respone.growth5, cultures), 2).ToString();
*/
           // respone.ebitda0 = Math.Round(Convert.ToDecimal(respone.ebitda0, cultures), 0).ToString();
            respone.ebitda1 = Math.Round(Convert.ToDecimal(respone.ebitda1, cultures), 0).ToString();
            respone.ebitda2 = Math.Round(Convert.ToDecimal(respone.ebitda2, cultures), 0).ToString();
            respone.ebitda3 = Math.Round(Convert.ToDecimal(respone.ebitda3, cultures), 0).ToString();
            respone.ebitda4 = Math.Round(Convert.ToDecimal(respone.ebitda4, cultures), 0).ToString();
            respone.ebitda5 = Math.Round(Convert.ToDecimal(respone.ebitda5, cultures), 0).ToString();

            respone.eper0 = Math.Round(Convert.ToDecimal(respone.eper0, cultures), 2).ToString();
            respone.eper1 = Math.Round(Convert.ToDecimal(respone.eper1, cultures), 2).ToString();
            respone.eper2 = Math.Round(Convert.ToDecimal(respone.eper2, cultures), 2).ToString();
            respone.eper3 = Math.Round(Convert.ToDecimal(respone.eper3, cultures), 2).ToString();
            respone.eper4 = Math.Round(Convert.ToDecimal(respone.eper4, cultures), 2).ToString();
            respone.eper5 = Math.Round(Convert.ToDecimal(respone.eper5, cultures), 2).ToString();

            respone.earnout1 = Math.Round(Convert.ToDecimal(respone.earnout1, cultures), 0).ToString();
            respone.earnout2 = Math.Round(Convert.ToDecimal(respone.earnout2, cultures), 0).ToString();
            respone.earnout3 = Math.Round(Convert.ToDecimal(respone.earnout3, cultures), 0).ToString();
            respone.earnout4 = Math.Round(Convert.ToDecimal(respone.earnout4, cultures), 0).ToString();
            respone.earnout5 = Math.Round(Convert.ToDecimal(respone.earnout5, cultures), 0).ToString();

            respone.remconspay1 = Math.Round(Convert.ToDecimal(respone.remconspay1, cultures), 0).ToString();
            respone.remconspay2 = Math.Round(Convert.ToDecimal(respone.remconspay2, cultures), 0).ToString();
            respone.remconspay3 = Math.Round(Convert.ToDecimal(respone.remconspay3, cultures), 0).ToString();
            respone.remconspay4 = Math.Round(Convert.ToDecimal(respone.remconspay4, cultures), 0).ToString();
            respone.remconspay5 = Math.Round(Convert.ToDecimal(respone.remconspay5, cultures), 0).ToString();

            respone.ebitdaearnout1 = Math.Round(Convert.ToDecimal(respone.ebitdaearnout1, cultures), 0).ToString();
            respone.ebitdaearnout2 = Math.Round(Convert.ToDecimal(respone.ebitdaearnout2, cultures), 0).ToString();
            respone.ebitdaearnout3 = Math.Round(Convert.ToDecimal(respone.ebitdaearnout3, cultures), 0).ToString();
            respone.ebitdaearnout4 = Math.Round(Convert.ToDecimal(respone.ebitdaearnout4, cultures), 0).ToString();
            respone.ebitdaearnout5 = Math.Round(Convert.ToDecimal(respone.ebitdaearnout5, cultures), 0).ToString();

            respone.depriciation1 = Math.Round(Convert.ToDecimal(respone.depriciation1, cultures), 0).ToString();
            respone.depriciation2 = Math.Round(Convert.ToDecimal(respone.depriciation2, cultures), 0).ToString();
            respone.depriciation3 = Math.Round(Convert.ToDecimal(respone.depriciation3, cultures), 0).ToString();
            respone.depriciation4 = Math.Round(Convert.ToDecimal(respone.depriciation4, cultures), 0).ToString();
            respone.depriciation5 = Math.Round(Convert.ToDecimal(respone.depriciation5, cultures), 0).ToString();

            respone.noncompamor1 = Math.Round(Convert.ToDecimal(respone.noncompamor1, cultures), 0).ToString();
            respone.noncompamor2 = Math.Round(Convert.ToDecimal(respone.noncompamor2, cultures), 0).ToString();
            respone.noncompamor3 = Math.Round(Convert.ToDecimal(respone.noncompamor3, cultures), 0).ToString();
            respone.noncompamor4 = Math.Round(Convert.ToDecimal(respone.noncompamor4, cultures), 0).ToString();
            respone.noncompamor5 = Math.Round(Convert.ToDecimal(respone.noncompamor5, cultures), 0).ToString();

            respone.pergoodwillamor1 = Math.Round(Convert.ToDecimal(respone.pergoodwillamor1, cultures), 0).ToString();
            respone.pergoodwillamor2 = Math.Round(Convert.ToDecimal(respone.pergoodwillamor2, cultures), 0).ToString();
            respone.pergoodwillamor3 = Math.Round(Convert.ToDecimal(respone.pergoodwillamor3, cultures), 0).ToString();
            respone.pergoodwillamor4 = Math.Round(Convert.ToDecimal(respone.pergoodwillamor4, cultures), 0).ToString();
            respone.pergoodwillamor5 = Math.Round(Convert.ToDecimal(respone.pergoodwillamor5, cultures), 0).ToString();

            respone.preconsamor1 = Math.Round(Convert.ToDecimal(respone.preconsamor1, cultures), 0).ToString();
            respone.preconsamor2 = Math.Round(Convert.ToDecimal(respone.preconsamor2, cultures), 0).ToString();
            respone.preconsamor3 = Math.Round(Convert.ToDecimal(respone.preconsamor3, cultures), 0).ToString();
            respone.preconsamor4 = Math.Round(Convert.ToDecimal(respone.preconsamor4, cultures), 0).ToString();
            respone.preconsamor5 = Math.Round(Convert.ToDecimal(respone.preconsamor5, cultures), 0).ToString();

            //respone.lessacquisition = Math.Round(Convert.ToDecimal(respone.lessacquisition, cultures), 0).ToString();
            // respone.lessacquisition = Math.Round(Convert.ToDecimal(respone.lessacquisition, cultures), 0).ToString();

            respone.acqcostamort1 = Math.Round(Convert.ToDecimal(respone.acqcostamort1, cultures), 0).ToString("#,####");
            respone.acqcostamort2 = Math.Round(Convert.ToDecimal(respone.acqcostamort2, cultures), 0).ToString("#,####");
            respone.acqcostamort3 = Math.Round(Convert.ToDecimal(respone.acqcostamort3, cultures), 0).ToString("#,####");
            respone.acqcostamort4 = Math.Round(Convert.ToDecimal(respone.acqcostamort4, cultures), 0).ToString("#,####");
            respone.acqcostamort5 = Math.Round(Convert.ToDecimal(respone.acqcostamort5, cultures), 0).ToString("#,####");

            respone.goodwillamorttax1 = Math.Round(Convert.ToDecimal(respone.goodwillamorttax1, cultures), 0).ToString("#,####");
            respone.goodwillamorttax2 = Math.Round(Convert.ToDecimal(respone.goodwillamorttax2, cultures), 0).ToString("#,####");
            respone.goodwillamorttax3 = Math.Round(Convert.ToDecimal(respone.goodwillamorttax3, cultures), 0).ToString("#,####");
            respone.goodwillamorttax4 = Math.Round(Convert.ToDecimal(respone.goodwillamorttax4, cultures), 0).ToString("#,####");
            respone.goodwillamorttax5 = Math.Round(Convert.ToDecimal(respone.goodwillamorttax5, cultures), 0).ToString("#,####");

            respone.totaldepandamort1 = Math.Round(Convert.ToDecimal(respone.totaldepandamort1, cultures), 0).ToString("#,####");
            respone.totaldepandamort2 = Math.Round(Convert.ToDecimal(respone.totaldepandamort2, cultures), 0).ToString("#,####");
            respone.totaldepandamort3 = Math.Round(Convert.ToDecimal(respone.totaldepandamort3, cultures), 0).ToString("#,####");
            respone.totaldepandamort4 = Math.Round(Convert.ToDecimal(respone.totaldepandamort4, cultures), 0).ToString("#,####");
            respone.totaldepandamort5 = Math.Round(Convert.ToDecimal(respone.totaldepandamort5, cultures), 0).ToString("#,####");

            respone.ebit1 = Math.Round(Convert.ToDecimal(respone.ebit1, cultures), 0).ToString("#,####");
            respone.ebit2 = Math.Round(Convert.ToDecimal(respone.ebit2, cultures), 0).ToString("#,####");
            respone.ebit3 = Math.Round(Convert.ToDecimal(respone.ebit3, cultures), 0).ToString("#,####");
            respone.ebit4 = Math.Round(Convert.ToDecimal(respone.ebit4, cultures), 0).ToString("#,####");
            respone.ebit5 = Math.Round(Convert.ToDecimal(respone.ebit5, cultures), 0).ToString("#,####");

            respone.intexprevolver1 = Math.Round(Convert.ToDecimal(respone.intexprevolver1, cultures), 0).ToString("#,####");
            respone.intexprevolver2 = Math.Round(Convert.ToDecimal(respone.intexprevolver2, cultures), 0).ToString("#,####");
            respone.intexprevolver3 = Math.Round(Convert.ToDecimal(respone.intexprevolver3, cultures), 0).ToString("#,####");
            respone.intexprevolver4 = Math.Round(Convert.ToDecimal(respone.intexprevolver4, cultures), 0).ToString("#,####");
            respone.intexprevolver5 = Math.Round(Convert.ToDecimal(respone.intexprevolver5, cultures), 0).ToString("#,####");

            respone.intexptermloan1 = Math.Round(Convert.ToDecimal(respone.intexptermloan1, cultures), 0).ToString("#,####");
            respone.intexptermloan2 = Math.Round(Convert.ToDecimal(respone.intexptermloan2, cultures), 0).ToString("#,####");
            respone.intexptermloan3 = Math.Round(Convert.ToDecimal(respone.intexptermloan3, cultures), 0).ToString("#,####");
            respone.intexptermloan4 = Math.Round(Convert.ToDecimal(respone.intexptermloan4, cultures), 0).ToString("#,####");
            respone.intexptermloan5 = Math.Round(Convert.ToDecimal(respone.intexptermloan5, cultures), 0).ToString("#,####");

            respone.intexpoveradvloan1 = Math.Round(Convert.ToDecimal(respone.intexpoveradvloan1, cultures), 0).ToString("#,####");
            respone.intexpoveradvloan2 = Math.Round(Convert.ToDecimal(respone.intexpoveradvloan2, cultures), 0).ToString("#,####");
            respone.intexpoveradvloan3 = Math.Round(Convert.ToDecimal(respone.intexpoveradvloan3, cultures), 0).ToString("#,####");
            respone.intexpoveradvloan4 = Math.Round(Convert.ToDecimal(respone.intexpoveradvloan4, cultures), 0).ToString("#,####");
            respone.intexpoveradvloan5 = Math.Round(Convert.ToDecimal(respone.intexpoveradvloan5, cultures), 0).ToString("#,####");

            respone.intexpmezzaninefinan1 = Math.Round(Convert.ToDecimal(respone.intexpmezzaninefinan1, cultures), 0).ToString("#,####");
            respone.intexpmezzaninefinan2 = Math.Round(Convert.ToDecimal(respone.intexpmezzaninefinan2, cultures), 0).ToString("#,####");
            respone.intexpmezzaninefinan3 = Math.Round(Convert.ToDecimal(respone.intexpmezzaninefinan3, cultures), 0).ToString("#,####");
            respone.intexpmezzaninefinan4 = Math.Round(Convert.ToDecimal(respone.intexpmezzaninefinan4, cultures), 0).ToString("#,####");
            respone.intexpmezzaninefinan5 = Math.Round(Convert.ToDecimal(respone.intexpmezzaninefinan5, cultures), 0).ToString("#,####");

            respone.intexpcapexloan1 = Math.Round(Convert.ToDecimal(respone.intexpcapexloan1, cultures), 0).ToString("#,####");
            respone.intexpcapexloan2 = Math.Round(Convert.ToDecimal(respone.intexpcapexloan2, cultures), 0).ToString("#,####");
            respone.intexpcapexloan3 = Math.Round(Convert.ToDecimal(respone.intexpcapexloan3, cultures), 0).ToString("#,####");
            respone.intexpcapexloan4 = Math.Round(Convert.ToDecimal(respone.intexpcapexloan4, cultures), 0).ToString("#,####");
            /*respone.intexpcapexloan5 = Math.Round(Convert.ToDecimal(respone.intexpcapexloan5, cultures), 0).ToString("#,####");

            respone.intexpgapnote1 = Math.Round(Convert.ToDecimal(respone.intexpgapnote1, cultures), 2).ToString();
            respone.intexpgapnote2 = Math.Round(Convert.ToDecimal(respone.intexpgapnote2, cultures), 2).ToString();
            respone.intexpgapnote3 = Math.Round(Convert.ToDecimal(respone.intexpgapnote3, cultures), 2).ToString();
            */respone.intexpgapnote4 = Math.Round(Convert.ToDecimal(respone.intexpgapnote4, cultures), 2).ToString();
            respone.intexpgapnote5 = Math.Round(Convert.ToDecimal(respone.intexpgapnote5, cultures), 2).ToString();

            respone.intexpgapballoonnote1 = Math.Round(Convert.ToDecimal(respone.intexpgapballoonnote1, cultures), 0).ToString("#,####");
            respone.intexpgapballoonnote2 = Math.Round(Convert.ToDecimal(respone.intexpgapballoonnote2, cultures), 0).ToString("#,####");
            respone.intexpgapballoonnote3 = Math.Round(Convert.ToDecimal(respone.intexpgapballoonnote3, cultures), 0).ToString("#,####");
            respone.intexpgapballoonnote4 = Math.Round(Convert.ToDecimal(respone.intexpgapballoonnote4, cultures), 0).ToString("#,####");
            respone.intexpgapballoonnote5 = Math.Round(Convert.ToDecimal(respone.intexpgapballoonnote5, cultures), 0).ToString("#,####");

            respone.intexpnoncompete1 = Math.Round(Convert.ToDecimal(respone.intexpnoncompete1, cultures), 0).ToString("#,####");
            respone.intexpnoncompete2 = Math.Round(Convert.ToDecimal(respone.intexpnoncompete2, cultures), 0).ToString("#,####");
            respone.intexpnoncompete3 = Math.Round(Convert.ToDecimal(respone.intexpnoncompete3, cultures), 0).ToString("#,####");
            respone.intexpnoncompete4 = Math.Round(Convert.ToDecimal(respone.intexpnoncompete4, cultures), 0).ToString("#,####");
            respone.intexpnoncompete5 = Math.Round(Convert.ToDecimal(respone.intexpnoncompete5, cultures), 0).ToString("#,####");

            respone.intexppersonalgoodwill1 = Math.Round(Convert.ToDecimal(respone.intexppersonalgoodwill1, cultures), 0).ToString("#,####");
            respone.intexppersonalgoodwill2 = Math.Round(Convert.ToDecimal(respone.intexppersonalgoodwill2, cultures), 0).ToString("#,####");
            respone.intexppersonalgoodwill3 = Math.Round(Convert.ToDecimal(respone.intexppersonalgoodwill3, cultures), 0).ToString("#,####");
            respone.intexppersonalgoodwill4 = Math.Round(Convert.ToDecimal(respone.intexppersonalgoodwill4, cultures), 0).ToString("#,####");
            respone.intexppersonalgoodwill5 = Math.Round(Convert.ToDecimal(respone.intexppersonalgoodwill5, cultures), 0).ToString("#,####");

            respone.intincomeoncash1 = Math.Round(Convert.ToDecimal(respone.intincomeoncash1, cultures), 0).ToString("#,####");
            respone.intincomeoncash2 = Math.Round(Convert.ToDecimal(respone.intincomeoncash2, cultures), 0).ToString("#,####");
            respone.intincomeoncash3 = Math.Round(Convert.ToDecimal(respone.intincomeoncash3, cultures), 0).ToString("#,####");
            respone.intincomeoncash4 = Math.Round(Convert.ToDecimal(respone.intincomeoncash4, cultures), 0).ToString("#,####");
            respone.intincomeoncash5 = Math.Round(Convert.ToDecimal(respone.intincomeoncash5, cultures), 0).ToString("#,####");

            respone.totalIntExpense1 = Math.Round(Convert.ToDecimal(respone.totalIntExpense1, cultures), 0).ToString("#,####");
            respone.totalIntExpense2 = Math.Round(Convert.ToDecimal(respone.totalIntExpense2, cultures), 0).ToString("#,####");
            respone.totalIntExpense3 = Math.Round(Convert.ToDecimal(respone.totalIntExpense3, cultures), 0).ToString("#,####");
            respone.totalIntExpense4 = Math.Round(Convert.ToDecimal(respone.totalIntExpense4, cultures), 0).ToString("#,####");
            respone.totalIntExpense5 = Math.Round(Convert.ToDecimal(respone.totalIntExpense5, cultures), 0).ToString("#,####");

            respone.taxableincome1 = Math.Round(Convert.ToDecimal(respone.taxableincome1, cultures), 0).ToString("#,####");
            respone.taxableincome2 = Math.Round(Convert.ToDecimal(respone.taxableincome2, cultures), 0).ToString("#,####");
            respone.taxableincome3 = Math.Round(Convert.ToDecimal(respone.taxableincome3, cultures), 0).ToString("#,####");
            respone.taxableincome4 = Math.Round(Convert.ToDecimal(respone.taxableincome4, cultures), 0).ToString("#,####");
            respone.taxableincome5 = Math.Round(Convert.ToDecimal(respone.taxableincome5, cultures), 0).ToString("#,####");

            respone.corpTaxesState1 = Math.Round(Convert.ToDecimal(respone.corpTaxesState1, cultures), 0).ToString("#,####");
            respone.corpTaxesState2 = Math.Round(Convert.ToDecimal(respone.corpTaxesState2, cultures), 0).ToString("#,####");
            respone.corpTaxesState3 = Math.Round(Convert.ToDecimal(respone.corpTaxesState3, cultures), 0).ToString("#,####");
            respone.corpTaxesState4 = Math.Round(Convert.ToDecimal(respone.corpTaxesState4, cultures), 0).ToString("#,####");
            respone.corpTaxesState5 = Math.Round(Convert.ToDecimal(respone.corpTaxesState5, cultures), 0).ToString("#,####");

            respone.corptexfederal1 = Math.Round(Convert.ToDecimal(respone.corptexfederal1, cultures), 0).ToString("#,####");
            respone.corptexfederal2 = Math.Round(Convert.ToDecimal(respone.corptexfederal2, cultures), 0).ToString("#,####");
            respone.corptexfederal3 = Math.Round(Convert.ToDecimal(respone.corptexfederal3, cultures), 0).ToString("#,####");
            respone.corptexfederal4 = Math.Round(Convert.ToDecimal(respone.corptexfederal4, cultures), 0).ToString("#,####");
            respone.corptexfederal5 = Math.Round(Convert.ToDecimal(respone.corptexfederal5, cultures), 0).ToString("#,####");

            respone.netincome1 = Math.Round(Convert.ToDecimal(respone.netincome1, cultures), 0).ToString("#,####");
            respone.netincome2 = Math.Round(Convert.ToDecimal(respone.netincome2, cultures), 0).ToString("#,####");
            respone.netincome3 = Math.Round(Convert.ToDecimal(respone.netincome3, cultures), 0).ToString("#,####");
            respone.netincome4 = Math.Round(Convert.ToDecimal(respone.netincome4, cultures), 0).ToString("#,####");
            respone.netincome5 = Math.Round(Convert.ToDecimal(respone.netincome5, cultures), 0).ToString("#,####");

            //Balance sheet

            respone.balancepurhaseCash = Math.Round(Convert.ToDecimal(respone.balancepurhaseCash, cultures),0).ToString("#,####");
            //respone.balanceopeningCash = Math.Round(Convert.ToDecimal(respone.balancepurhaseCash, cultures),0).ToString("#,####");
            respone.balanceCashyear1 = Math.Round(Convert.ToDecimal(respone.balanceCashyear1, cultures), 0).ToString("#,####");
            respone.balanceCashyear2 = Math.Round(Convert.ToDecimal(respone.balanceCashyear2, cultures), 0).ToString("#,####");
            respone.balanceCashyear3 = Math.Round(Convert.ToDecimal(respone.balanceCashyear3, cultures), 0).ToString("#,####");
            respone.balanceCashyear4 = Math.Round(Convert.ToDecimal(respone.balanceCashyear4, cultures), 0).ToString("#,####");
            respone.balanceCashyear5 = Math.Round(Convert.ToDecimal(respone.balanceCashyear5, cultures), 0).ToString("#,####");



            respone.balancepurchaseAr = Math.Round(Convert.ToDecimal(respone.balancepurchaseAr, cultures), 0).ToString("#,####");
            respone.balanceopeningAr = Math.Round(Convert.ToDecimal(respone.balanceopeningAr, cultures), 0).ToString("#,####");
            respone.balanceAryear1 = Math.Round(Convert.ToDecimal(respone.balanceAryear1, cultures), 0).ToString("#,####");
            respone.balanceAryear2 = Math.Round(Convert.ToDecimal(respone.balanceAryear2, cultures), 0).ToString("#,####");
            respone.balanceAryear3 = Math.Round(Convert.ToDecimal(respone.balanceAryear3, cultures), 0).ToString("#,####");
            respone.balanceAryear4 = Math.Round(Convert.ToDecimal(respone.balanceAryear4, cultures), 0).ToString("#,####");
            respone.balanceAryear5 = Math.Round(Convert.ToDecimal(respone.balanceAryear5, cultures), 0).ToString("#,####");



            respone.balancepurchaseInventory = Math.Round(Convert.ToDecimal(respone.balancepurchaseInventory, cultures), 0).ToString("#,####");
            respone.balanceopeningInventory = Math.Round(Convert.ToDecimal(respone.balanceopeningInventory, cultures), 0).ToString("#,####");
            respone.balanceInventoryyear1 = Math.Round(Convert.ToDecimal(respone.balanceInventoryyear1, cultures), 0).ToString("#,####");
            respone.balanceInventoryyear2 = Math.Round(Convert.ToDecimal(respone.balanceInventoryyear2, cultures), 0).ToString("#,####");
            respone.balanceInventoryyear3 = Math.Round(Convert.ToDecimal(respone.balanceInventoryyear3, cultures), 0).ToString("#,####");
            respone.balanceInventoryyear4 = Math.Round(Convert.ToDecimal(respone.balanceInventoryyear4, cultures), 0).ToString("#,####");
            respone.balanceInventoryyear5 = Math.Round(Convert.ToDecimal(respone.balanceInventoryyear5, cultures), 0).ToString("#,####");



            respone.balancepurchasemiscAssets = Math.Round(Convert.ToDecimal(respone.balancepurchasemiscAssets, cultures), 0).ToString("#,####");
            respone.balanceopeningmiscAssets = Math.Round(Convert.ToDecimal(respone.balanceopeningmiscAssets, cultures), 0).ToString("#,####");
            respone.balancemiscAssetsyear1 = Math.Round(Convert.ToDecimal(respone.balancemiscAssetsyear1, cultures), 0).ToString("#,####");
            respone.balancemiscAssetsyear2 = Math.Round(Convert.ToDecimal(respone.balancemiscAssetsyear2, cultures), 0).ToString("#,####");
            respone.balancemiscAssetsyear3 = Math.Round(Convert.ToDecimal(respone.balancemiscAssetsyear3, cultures), 0).ToString("#,####");
            respone.balancemiscAssetsyear4 = Math.Round(Convert.ToDecimal(respone.balancemiscAssetsyear4, cultures), 0).ToString("#,####");
            respone.balancemiscAssetsyear5 = Math.Round(Convert.ToDecimal(respone.balancemiscAssetsyear5, cultures), 0).ToString("#,####");


            respone.balancepurchaseFixedAssOld = Math.Round(Convert.ToDecimal(respone.balancepurchaseFixedAssOld, cultures), 0).ToString("#,####");
            respone.balanceopeningFixedAssOld = Math.Round(Convert.ToDecimal(respone.balanceopeningFixedAssOld, cultures), 0).ToString("#,####");
            respone.balanceFixedAssOldyear1 = Math.Round(Convert.ToDecimal(respone.balanceFixedAssOldyear1, cultures), 0).ToString("#,####");
            respone.balanceFixedAssOldyear2 = Math.Round(Convert.ToDecimal(respone.balanceFixedAssOldyear2, cultures), 0).ToString("#,####");
            respone.balanceFixedAssOldyear3 = Math.Round(Convert.ToDecimal(respone.balanceFixedAssOldyear3, cultures), 0).ToString("#,####");
            respone.balanceFixedAssOldyear4 = Math.Round(Convert.ToDecimal(respone.balanceFixedAssOldyear4, cultures), 0).ToString("#,####");
            respone.balanceFixedAssOldyear5 = Math.Round(Convert.ToDecimal(respone.balanceFixedAssOldyear5, cultures), 0).ToString("#,####");



            respone.balanceopeningADold = Math.Round(Convert.ToDecimal(respone.balanceopeningADold, cultures), 0).ToString("#,####");
            respone.balanceADoldYear1 = Math.Round(Convert.ToDecimal(respone.balanceADoldYear1, cultures), 0).ToString("#,####");
            respone.balanceADoldYear2 = Math.Round(Convert.ToDecimal(respone.balanceADoldYear2, cultures), 0).ToString("#,####");
            respone.balanceADoldYear3 = Math.Round(Convert.ToDecimal(respone.balanceADoldYear3, cultures), 0).ToString("#,####");
            respone.balanceADoldYear4 = Math.Round(Convert.ToDecimal(respone.balanceADoldYear4, cultures), 0).ToString("#,####");
            respone.balanceADoldYear5 = Math.Round(Convert.ToDecimal(respone.balanceADoldYear5, cultures), 0).ToString("#,####");


            respone.balanceopenNewFixedAssets = Math.Round(Convert.ToDecimal(respone.balanceopenNewFixedAssets, cultures), 0).ToString("#,####");
            respone.balanceNewFixedAssYear1 = Math.Round(Convert.ToDecimal(respone.balanceNewFixedAssYear1, cultures), 0).ToString("#,####");
            respone.balanceNewFixedAssYear2 = Math.Round(Convert.ToDecimal(respone.balanceNewFixedAssYear2, cultures), 0).ToString("#,####");
            respone.balanceNewFixedAssYear3 = Math.Round(Convert.ToDecimal(respone.balanceNewFixedAssYear3, cultures), 0).ToString("#,####");
            respone.balanceNewFixedAssYear4 = Math.Round(Convert.ToDecimal(respone.balanceNewFixedAssYear4, cultures), 0).ToString("#,####");
            respone.balanceNewFixedAssYear5 = Math.Round(Convert.ToDecimal(respone.balanceNewFixedAssYear5, cultures), 0).ToString("#,####");


            respone.balanceopenADNewFixedAssets = Math.Round(Convert.ToDecimal(respone.balanceopenADNewFixedAssets, cultures), 0).ToString("#,####");
            respone.balanceADNewFixedAssyear1 = Math.Round(Convert.ToDecimal(respone.balanceADNewFixedAssyear1, cultures), 0).ToString("#,####");
            respone.balanceADNewFixedAssyear2 = Math.Round(Convert.ToDecimal(respone.balanceADNewFixedAssyear2, cultures), 0).ToString("#,####");
            respone.balanceADNewFixedAssyear3 = Math.Round(Convert.ToDecimal(respone.balanceADNewFixedAssyear3, cultures), 0).ToString("#,####");
            respone.balanceADNewFixedAssyear4 = Math.Round(Convert.ToDecimal(respone.balanceADNewFixedAssyear4, cultures), 0).ToString("#,####");
            respone.balanceADNewFixedAssyear5 = Math.Round(Convert.ToDecimal(respone.balanceADNewFixedAssyear5, cultures), 0).ToString("#,####");

            respone.balanceOpenAcquisitionExp = Math.Round(Convert.ToDecimal(respone.balanceOpenAcquisitionExp, cultures), 0).ToString("#,####");
            respone.balanceAcquisitionYear1 = Math.Round(Convert.ToDecimal(respone.balanceAcquisitionYear1, cultures), 0).ToString("#,####");
            respone.balanceAcquisitionYear2 = Math.Round(Convert.ToDecimal(respone.balanceAcquisitionYear2, cultures), 0).ToString("#,####");
            respone.balanceAcquisitionYear3 = Math.Round(Convert.ToDecimal(respone.balanceAcquisitionYear3, cultures), 0).ToString("#,####");
            respone.balanceAcquisitionYear4 = Math.Round(Convert.ToDecimal(respone.balanceAcquisitionYear4, cultures), 0).ToString("#,####");
            respone.balanceAcquisitionYear5 = Math.Round(Convert.ToDecimal(respone.balanceAcquisitionYear5, cultures), 0).ToString("#,####");


            respone.balanceOpenNonCompete = Math.Round(Convert.ToDecimal(respone.balanceOpenNonCompete, cultures), 0).ToString("#,####");
            respone.balanceNonCompeteYear1 = Math.Round(Convert.ToDecimal(respone.balanceNonCompeteYear1, cultures), 0).ToString("#,####");
            respone.balanceNonCompeteYear2 = Math.Round(Convert.ToDecimal(respone.balanceNonCompeteYear2, cultures), 0).ToString("#,####");
            respone.balanceNonCompeteYear3 = Math.Round(Convert.ToDecimal(respone.balanceNonCompeteYear3, cultures), 0).ToString("#,####");
            respone.balanceNonCompeteYear4 = Math.Round(Convert.ToDecimal(respone.balanceNonCompeteYear4, cultures), 0).ToString("#,####");
            respone.balanceNonCompeteYear5 = Math.Round(Convert.ToDecimal(respone.balanceNonCompeteYear5, cultures), 0).ToString("#,####");


            respone.balanceopenPersonalGoodwill = Math.Round(Convert.ToDecimal(respone.balanceopenPersonalGoodwill, cultures), 0).ToString("#,####");
            respone.balancePersonalGoodwillYear1 = Math.Round(Convert.ToDecimal(respone.balancePersonalGoodwillYear1, cultures), 0).ToString("#,####");
            respone.balancePersonalGoodwillYear2 = Math.Round(Convert.ToDecimal(respone.balancePersonalGoodwillYear2, cultures), 0).ToString("#,####");
            respone.balancePersonalGoodwillYear3 = Math.Round(Convert.ToDecimal(respone.balancePersonalGoodwillYear3, cultures), 0).ToString("#,####");
            respone.balancePersonalGoodwillYear4 = Math.Round(Convert.ToDecimal(respone.balancePersonalGoodwillYear4, cultures), 0).ToString("#,####");
            respone.balancePersonalGoodwillYear5 = Math.Round(Convert.ToDecimal(respone.balancePersonalGoodwillYear5, cultures), 0).ToString("#,####");

            respone.balanceOpenRemCons1 = Math.Round(Convert.ToDecimal(respone.balanceOpenRemCons1, cultures), 0).ToString("#,####");
            respone.balanceRemConsYear1 = Math.Round(Convert.ToDecimal(respone.balanceRemConsYear1, cultures), 0).ToString("#,####");
            respone.balanceRemConsYear2 = Math.Round(Convert.ToDecimal(respone.balanceRemConsYear2, cultures), 0).ToString("#,####");
            respone.balanceRemConsYear3 = Math.Round(Convert.ToDecimal(respone.balanceRemConsYear3, cultures), 0).ToString("#,####");
            respone.balanceRemConsYear4 = Math.Round(Convert.ToDecimal(respone.balanceRemConsYear4, cultures), 0).ToString("#,####");
            respone.balanceRemConsYear5 = Math.Round(Convert.ToDecimal(respone.balanceRemConsYear5, cultures), 0).ToString("#,####");

            respone.balanceOpenPreCons = Math.Round(Convert.ToDecimal(respone.balanceOpenPreCons, cultures), 0).ToString("#,####");
            respone.balancePrepaidConsYear1 = Math.Round(Convert.ToDecimal(respone.balancePrepaidConsYear1, cultures), 0).ToString("#,####");
            respone.balancePrepaidConsYear2 = Math.Round(Convert.ToDecimal(respone.balancePrepaidConsYear2, cultures), 0).ToString("#,####");
            respone.balancePrepaidConsYear3 = Math.Round(Convert.ToDecimal(respone.balancePrepaidConsYear3, cultures), 0).ToString("#,####");
            respone.balancePrepaidConsYear4 = Math.Round(Convert.ToDecimal(respone.balancePrepaidConsYear4, cultures), 0).ToString("#,####");
            respone.balancePrepaidConsYear5 = Math.Round(Convert.ToDecimal(respone.balancePrepaidConsYear5, cultures), 0).ToString("#,####");

            respone.balanceopenInvREntity = Math.Round(Convert.ToDecimal(respone.balanceopenInvREntity, cultures), 0).ToString("#,####");
            respone.balanceInvREntityYear1 = Math.Round(Convert.ToDecimal(respone.balanceInvREntityYear1, cultures), 0).ToString("#,####");
            respone.balanceInvREntityYear2 = Math.Round(Convert.ToDecimal(respone.balanceInvREntityYear2, cultures), 0).ToString("#,####");
            respone.balanceInvREntityYear3 = Math.Round(Convert.ToDecimal(respone.balanceInvREntityYear3, cultures), 0).ToString("#,####");
            respone.balanceInvREntityYear4 = Math.Round(Convert.ToDecimal(respone.balanceInvREntityYear4, cultures), 0).ToString("#,####");
            respone.balanceInvREntityYear5 = Math.Round(Convert.ToDecimal(respone.balanceInvREntityYear5, cultures), 0).ToString("#,####");

            respone.balancepurchGoodwill = Math.Round(Convert.ToDecimal(respone.balancepurchGoodwill, cultures), 0).ToString("#,####");
            respone.balanceopenGoodwillRes = Math.Round(Convert.ToDecimal(respone.balanceopenGoodwillRes, cultures), 0).ToString("#,####");
            respone.balanceGoodwillYear1 = Math.Round(Convert.ToDecimal(respone.balanceGoodwillYear1, cultures), 0).ToString("#,####");
            respone.balanceGoodwillYear2 = Math.Round(Convert.ToDecimal(respone.balanceGoodwillYear2, cultures), 0).ToString("#,####");
            respone.balanceGoodwillYear3 = Math.Round(Convert.ToDecimal(respone.balanceGoodwillYear3, cultures), 0).ToString("#,####");
            respone.balanceGoodwillYear4 = Math.Round(Convert.ToDecimal(respone.balanceGoodwillYear4, cultures), 0).ToString("#,####");
            respone.balanceGoodwillYear5 = Math.Round(Convert.ToDecimal(respone.balanceGoodwillYear5, cultures), 0).ToString("#,####");

            respone.balancepurchTotalAssets = Math.Round(Convert.ToDecimal(respone.balancepurchTotalAssets, cultures), 0).ToString("#,####");
            respone.balanceopenTotalAssets = Math.Round(Convert.ToDecimal(respone.balanceopenTotalAssets, cultures), 0).ToString("#,####");
            respone.balanceTotalAssetsYear1 = Math.Round(Convert.ToDecimal(respone.balanceTotalAssetsYear1, cultures), 0).ToString("#,####");
            respone.balanceTotalAssetsYear2 = Math.Round(Convert.ToDecimal(respone.balanceTotalAssetsYear2, cultures), 0).ToString("#,####");
            respone.balanceTotalAssetsYear3 = Math.Round(Convert.ToDecimal(respone.balanceTotalAssetsYear3, cultures), 0).ToString("#,####");
            respone.balanceTotalAssetsYear4 = Math.Round(Convert.ToDecimal(respone.balanceTotalAssetsYear4, cultures), 0).ToString("#,####");
            respone.balanceTotalAssetsYear5 = Math.Round(Convert.ToDecimal(respone.balanceTotalAssetsYear5, cultures), 0).ToString("#,####");

            respone.balancepurchAPaccured = Math.Round(Convert.ToDecimal(respone.balancepurchAPaccured, cultures), 0).ToString("#,####");
            respone.balanceopenARAccured = Math.Round(Convert.ToDecimal(respone.balanceopenARAccured, cultures), 0).ToString("#,####");
            respone.balanceARAccuredYear1 = Math.Round(Convert.ToDecimal(respone.balanceARAccuredYear1, cultures), 0).ToString("#,####");
            respone.balanceARAccuredYear2 = Math.Round(Convert.ToDecimal(respone.balanceARAccuredYear2, cultures), 0).ToString("#,####");
            respone.balanceARAccuredYear3 = Math.Round(Convert.ToDecimal(respone.balanceARAccuredYear3, cultures), 0).ToString("#,####");
            respone.balanceARAccuredYear4 = Math.Round(Convert.ToDecimal(respone.balanceARAccuredYear4, cultures), 0).ToString("#,####");
            respone.balanceARAccuredYear5 = Math.Round(Convert.ToDecimal(respone.balanceARAccuredYear5, cultures), 0).ToString("#,####");

            respone.balancepurchOtherMiscLiab = Math.Round(Convert.ToDecimal(respone.balancepurchOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceopenOtherMiscLiab = Math.Round(Convert.ToDecimal(respone.balanceopenOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceOtherMiscLiabYear1 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear1, cultures), 0).ToString("#,####");
            respone.balanceOtherMiscLiabYear2 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear2, cultures), 0).ToString("#,####");
            respone.balanceOtherMiscLiabYear3 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear3, cultures), 0).ToString("#,####");
            respone.balanceOtherMiscLiabYear4 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear4, cultures), 0).ToString("#,####");
            respone.balanceOtherMiscLiabYear5 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear5, cultures), 0).ToString("#,####");

            respone.balancepurchNonOperLiab = Math.Round(Convert.ToDecimal(respone.balancepurchNonOperLiab, cultures), 0).ToString("#,####");
            respone.balanceopenNonOperLiab = Math.Round(Convert.ToDecimal(respone.balanceopenNonOperLiab, cultures), 0).ToString("#,####");
            respone.balanceNonOperLiabYear1 = Math.Round(Convert.ToDecimal(respone.balanceNonOperLiabYear1, cultures), 0).ToString("#,####");
            respone.balanceNonOperLiabYear2 = Math.Round(Convert.ToDecimal(respone.balanceNonOperLiabYear2, cultures), 0).ToString("#,####");
            respone.balanceNonOperLiabYear3 = Math.Round(Convert.ToDecimal(respone.balanceNonOperLiabYear3, cultures), 0).ToString("#,####");
            respone.balanceNonOperLiabYear4 = Math.Round(Convert.ToDecimal(respone.balanceNonOperLiabYear4, cultures), 0).ToString("#,####");
            respone.balanceNonOperLiabYear5 = Math.Round(Convert.ToDecimal(respone.balanceNonOperLiabYear5, cultures), 0).ToString("#,####");

            respone.balanceopenRevolver = Math.Round(Convert.ToDecimal(respone.balanceopenRevolver, cultures), 2).ToString();
            respone.balanceRevolverYear1 = Math.Round(Convert.ToDecimal(respone.balanceRevolverYear1, cultures), 2).ToString();
            respone.balanceRevolverYear2 = Math.Round(Convert.ToDecimal(respone.balanceRevolverYear2, cultures), 2).ToString();
            respone.balanceRevolverYear3 = Math.Round(Convert.ToDecimal(respone.balanceRevolverYear3, cultures), 2).ToString();
            respone.balanceRevolverYear4 = Math.Round(Convert.ToDecimal(respone.balanceRevolverYear4, cultures), 2).ToString();
            respone.balanceRevolverYear5 = Math.Round(Convert.ToDecimal(respone.balanceRevolverYear5, cultures), 2).ToString();

            respone.balanceopenTermLoan = Math.Round(Convert.ToDecimal(respone.balanceopenTermLoan, cultures), 2).ToString("#,####");
            respone.balanceTermLoanYear1 = Math.Round(Convert.ToDecimal(respone.balanceTermLoanYear1, cultures), 2).ToString("#,####");
            respone.balanceTermLoanYear2 = Math.Round(Convert.ToDecimal(respone.balanceTermLoanYear2, cultures), 2).ToString("#,####");
            respone.balanceTermLoanYear3 = Math.Round(Convert.ToDecimal(respone.balanceTermLoanYear3, cultures), 2).ToString("#,####");
            respone.balanceTermLoanYear4 = Math.Round(Convert.ToDecimal(respone.balanceTermLoanYear4, cultures), 2).ToString("#,####");
            /*respone.balanceTermLoanYear5 = Math.Round(Convert.ToDecimal(respone.balanceTermLoanYear5, cultures), 2).ToString();

            respone.balanceopenOverAdvLoan = Math.Round(Convert.ToDecimal(respone.balancepurchOtherMiscLiab, cultures), 2).ToString();
            respone.balanceOverAdvLoan1 = Math.Round(Convert.ToDecimal(respone.balanceopenOtherMiscLiab, cultures), 2).ToString();
            respone.balanceOverAdvLoan2 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear1, cultures), 2).ToString();
            *//*respone.balanceOverAdvLoan3 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear2, cultures), 2).ToString();
            respone.balanceOverAdvLoan4 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear3, cultures), 2).ToString();
            respone.balanceOverAdvLoan5 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear4, cultures), 2).ToString();
*/
           /* respone.balanceopenMezzFinancing = Math.Round(Convert.ToDecimal(respone.balancepurchOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceMezzFinancingYear1 = Math.Round(Convert.ToDecimal(respone.balanceopenOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceMezzFinancingYear2 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear1, cultures), 0).ToString("#,####");
            respone.balanceMezzFinancingYear3 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear2, cultures), 0).ToString("#,####");
            respone.balanceMezzFinancingYear4 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear3, cultures), 0).ToString("#,####");
            respone.balanceMezzFinancingYear5 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear4, cultures), 0).ToString("#,####");
*/
           /* respone.balanceopenGapNote = Math.Round(Convert.ToDecimal(respone.balancepurchOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceGapNoteYear1 = Math.Round(Convert.ToDecimal(respone.balanceopenOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceGapNoteYear2 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear1, cultures), 0).ToString("#,####");
            respone.balanceGapNoteYear3 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear2, cultures), 0).ToString("#,####");
            respone.balanceGapNoteYear4 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear3, cultures), 0).ToString("#,####");
            respone.balanceGapNoteYear5 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear4, cultures), 0).ToString("#,####");
*/
           /* respone.balanceopenGapBallonNote = Math.Round(Convert.ToDecimal(respone.balancepurchOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceGapBallonNoteYear1 = Math.Round(Convert.ToDecimal(respone.balanceopenOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceGapBallonNoteYear2 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear1, cultures), 0).ToString("#,####");
            respone.balanceGapBallonNoteYear3 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear2, cultures), 0).ToString("#,####");
            respone.balanceGapBallonNoteYear4 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear3, cultures), 0).ToString("#,####");
            respone.balanceGapBallonNoteYear5 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear4, cultures), 0).ToString("#,####");
*/
          /*  respone.balanceopenCapExLoan = Math.Round(Convert.ToDecimal(respone.balancepurchOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceCapExLoanYear1 = Math.Round(Convert.ToDecimal(respone.balanceopenOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceCapExLoanYear2 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear1, cultures), 0).ToString("#,####");
            respone.balanceCapExLoanYear3 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear2, cultures), 0).ToString("#,####");
            respone.balanceCapExLoanYear4 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear3, cultures), 0).ToString("#,####");
            respone.balanceCapExLoanYear5 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear4, cultures), 0).ToString("#,####");

            respone.balanceopenRemNonCompPayment = Math.Round(Convert.ToDecimal(respone.balancepurchOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceRemNonComPaymentYear1 = Math.Round(Convert.ToDecimal(respone.balanceopenOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceRemNonComPaymentYear2 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear1, cultures), 0).ToString("#,####");
            respone.balanceRemNonComPaymentYear3 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear2, cultures), 0).ToString("#,####");
            respone.balanceRemNonComPaymentYear4 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear3, cultures), 0).ToString("#,####");
            respone.balanceRemNonComPaymentYear5 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear4, cultures), 0).ToString("#,####");

            respone.balanceopenRemPersonalGWpay = Math.Round(Convert.ToDecimal(respone.balancepurchOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceRemPersonalGWpayYear1 = Math.Round(Convert.ToDecimal(respone.balanceopenOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceRemPersonalGWpayYear2 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear1, cultures), 0).ToString("#,####");
            respone.balanceRemPersonalGWpayYear3 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear2, cultures), 0).ToString("#,####");
            respone.balanceRemPersonalGWpayYear4 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear3, cultures), 0).ToString("#,####");
            respone.balanceRemPersonalGWpayYear5 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear4, cultures), 0).ToString("#,####");

            respone.balanceopenRemConsPay = Math.Round(Convert.ToDecimal(respone.balancepurchOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceRemConssPayYear1 = Math.Round(Convert.ToDecimal(respone.balanceopenOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceRemConssPayYear2 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear1, cultures), 0).ToString("#,####");
            respone.balanceRemConssPayYear3 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear2, cultures), 0).ToString("#,####");
            respone.balanceRemConssPayYear4 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear3, cultures), 0).ToString("#,####");
            respone.balanceRemConssPayYear5 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear4, cultures), 0).ToString("#,####");

            respone.balanceopenNonOperatingLiab = Math.Round(Convert.ToDecimal(respone.balancepurchOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceNonOperatingLiabYear1 = Math.Round(Convert.ToDecimal(respone.balanceopenOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceNonOperatingLiabYear2 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear1, cultures), 0).ToString("#,####");
            respone.balanceNonOperatingLiabYear3 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear2, cultures), 0).ToString("#,####");
            respone.balanceNonOperatingLiabYear4 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear3, cultures), 0).ToString("#,####");
            respone.balanceNonOperatingLiabYear5 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear4, cultures), 0).ToString("#,####");

            respone.balancepurchTotalLiab = Math.Round(Convert.ToDecimal(respone.balancepurchOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceopenTotalLiab = Math.Round(Convert.ToDecimal(respone.balanceopenOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceTotalLiabYear1 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear1, cultures), 0).ToString("#,####");
            respone.balanceTotalLiabYear2 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear2, cultures), 0).ToString("#,####");
            respone.balanceTotalLiabYear3 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear3, cultures), 0).ToString("#,####");
            respone.balanceTotalLiabYear4 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear4, cultures), 0).ToString("#,####");
            respone.balanceTotalLiabYear5 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear5, cultures), 0).ToString("#,####");

            respone.balancepurchRetainedEarning = Math.Round(Convert.ToDecimal(respone.balancepurchOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceopenRetaiedEarning = Math.Round(Convert.ToDecimal(respone.balanceopenOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceRetainedEarningYr1 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear1, cultures), 0).ToString("#,####");
            respone.balanceRetainedEarningYr2 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear2, cultures), 0).ToString("#,####");
            respone.balanceRetainedEarningYr3 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear3, cultures), 0).ToString("#,####");
            respone.balanceRetainedEarningYr4 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear4, cultures), 0).ToString("#,####");
            respone.balanceRetainedEarningYr5 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear5, cultures), 0).ToString("#,####");

            respone.balanceopenAddCapital = Math.Round(Convert.ToDecimal(respone.balancepurchOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceAddCapitalYear1 = Math.Round(Convert.ToDecimal(respone.balanceopenOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceAddCapitalYear2 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear1, cultures), 0).ToString("#,####");
            respone.balanceAddCapitalYear3 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear2, cultures), 0).ToString("#,####");
            respone.balanceAddCapitalYear4 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear3, cultures), 0).ToString("#,####");
            respone.balanceAddCapitalYear5 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear4, cultures), 0).ToString("#,####");
*/
           /* respone.balanceopenDisTax = Math.Round(Convert.ToDecimal(respone.balancepurchOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceDisTaxYear1 = Math.Round(Convert.ToDecimal(respone.balanceopenOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceDisTaxYear2 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear1, cultures), 0).ToString("#,####");
            respone.balanceDisTaxYear3 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear2, cultures), 0).ToString("#,####");
            respone.balanceDisTaxYear4 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear3, cultures), 0).ToString("#,####");
            respone.balanceDisTaxYear5 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear4, cultures), 0).ToString("#,####");

            respone.balanceopenDividend = Math.Round(Convert.ToDecimal(respone.balancepurchOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceDividendYear1 = Math.Round(Convert.ToDecimal(respone.balanceopenOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceDividendYear2 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear1, cultures), 0).ToString("#,####");
            respone.balanceDividendYear3 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear2, cultures), 0).ToString("#,####");
            respone.balanceDividendYear4 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear3, cultures), 0).ToString("#,####");
            respone.balanceDividendYear5 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear4, cultures), 0).ToString("#,####");

            respone.balancepurchCommonStk = Math.Round(Convert.ToDecimal(respone.balancepurchOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceopenCommonStk = Math.Round(Convert.ToDecimal(respone.balanceopenOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceCommonStkYear1 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear1, cultures), 0).ToString("#,####");
            respone.balanceCommonStkYear2 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear2, cultures), 0).ToString("#,####");
            respone.balanceCommonStkYear3 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear3, cultures), 0).ToString("#,####");
            respone.balanceCommonStkYear4 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear4, cultures), 0).ToString("#,####");
            respone.balanceCommonStkYear5 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear5, cultures), 0).ToString("#,####");

            respone.balancepurchTotalEquity = Math.Round(Convert.ToDecimal(respone.balancepurchOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceopenTotalEquity = Math.Round(Convert.ToDecimal(respone.balanceopenOtherMiscLiab, cultures), 0).ToString("#,####");
            respone.balanceTotalEquYear1 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear1, cultures), 0).ToString("#,####");
            respone.balanceTotalEquYear2 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear2, cultures), 0).ToString("#,####");
            respone.balanceTotalEquYear3 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear3, cultures), 0).ToString("#,####");
            respone.balanceTotalEquYear4 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear4, cultures), 0).ToString("#,####");
            respone.balanceTotalEquYear5 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear5, cultures), 0).ToString("#,####");

            respone.balancepurchTotalLiabEqu = Math.Round(Convert.ToDecimal(respone.balancepurchOtherMiscLiab, cultures), 2).ToString();
            respone.balanceopenTotalLiabEqu = Math.Round(Convert.ToDecimal(respone.balanceopenOtherMiscLiab, cultures), 2).ToString();
            respone.balanceTotalLiabEquYear1 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear1, cultures), 2).ToString();
            respone.balanceTotalLiabEquYear2 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear2, cultures), 2).ToString();
            respone.balanceTotalLiabEquYear3 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear3, cultures), 2).ToString();
            respone.balanceTotalLiabEquYear4 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear4, cultures), 2).ToString();
            respone.balanceTotalLiabEquYear5 = Math.Round(Convert.ToDecimal(respone.balanceOtherMiscLiabYear5, cultures), 2).ToString();

            //FREE CASH FLOW

            respone.freeNetIncomeYr1 = Math.Round(Convert.ToDecimal(respone.freeNetIncomeYr1, cultures), 0).ToString("#,####");
            respone.freeNetIncomeYr2 = Math.Round(Convert.ToDecimal(respone.freeNetIncomeYr2, cultures), 0).ToString("#,####");
            respone.freeNetIncomeYr3 = Math.Round(Convert.ToDecimal(respone.freeNetIncomeYr3, cultures), 0).ToString("#,####");
            respone.freeNetIncomeYr4 = Math.Round(Convert.ToDecimal(respone.freeNetIncomeYr4, cultures), 0).ToString("#,####");
            respone.freeNetIncomeYr5 = Math.Round(Convert.ToDecimal(respone.freeNetIncomeYr5, cultures), 0).ToString("#,####");

            respone.freedeprYr1 = Math.Round(Convert.ToDecimal(respone.freedeprYr1, cultures), 0).ToString("#,####");
            respone.freedeprYr2 = Math.Round(Convert.ToDecimal(respone.freedeprYr2, cultures), 0).ToString("#,####");
            respone.freedeprYr3 = Math.Round(Convert.ToDecimal(respone.freedeprYr3, cultures), 0).ToString("#,####");
            respone.freedeprYr4 = Math.Round(Convert.ToDecimal(respone.freedeprYr4, cultures), 0).ToString("#,####");
            respone.freedeprYr5 = Math.Round(Convert.ToDecimal(respone.freedeprYr5, cultures), 0).ToString("#,####");

            respone.freeNonCompAmortYr1 = Math.Round(Convert.ToDecimal(respone.freeNonCompAmortYr1, cultures), 0).ToString("#,####");
            respone.freeNonCompAmortYr2 = Math.Round(Convert.ToDecimal(respone.freeNonCompAmortYr2, cultures), 0).ToString("#,####");
            respone.freeNonCompAmortYr3 = Math.Round(Convert.ToDecimal(respone.freeNonCompAmortYr3, cultures), 0).ToString("#,####");
            respone.freeNonCompAmortYr4 = Math.Round(Convert.ToDecimal(respone.freeNonCompAmortYr4, cultures), 0).ToString("#,####");
            respone.freeNonCompAmortYr5 = Math.Round(Convert.ToDecimal(respone.freeNonCompAmortYr5, cultures), 0).ToString("#,####");

            respone.freePerGoodwillAmortYr1 = Math.Round(Convert.ToDecimal(respone.freePerGoodwillAmortYr1, cultures), 0).ToString("#,####");
            respone.freePerGoodwillAmortYr2 = Math.Round(Convert.ToDecimal(respone.freePerGoodwillAmortYr2, cultures), 0).ToString("#,####");
            respone.freePerGoodwillAmortYr3 = Math.Round(Convert.ToDecimal(respone.freePerGoodwillAmortYr3, cultures), 0).ToString("#,####");
            respone.freePerGoodwillAmortYr4 = Math.Round(Convert.ToDecimal(respone.freePerGoodwillAmortYr4, cultures), 0).ToString("#,####");
            respone.freePerGoodwillAmortYr5 = Math.Round(Convert.ToDecimal(respone.freePerGoodwillAmortYr5, cultures), 0).ToString("#,####");

            respone.freePreConsAmortYr1 = Math.Round(Convert.ToDecimal(respone.freePreConsAmortYr1, cultures), 0).ToString("#,####");
            respone.freePreConsAmortYr2 = Math.Round(Convert.ToDecimal(respone.freePreConsAmortYr2, cultures), 0).ToString("#,####");
            respone.freePreConsAmortYr3 = Math.Round(Convert.ToDecimal(respone.freePreConsAmortYr3, cultures), 0).ToString("#,####");
            respone.freePreConsAmortYr4 = Math.Round(Convert.ToDecimal(respone.freePreConsAmortYr4, cultures), 0).ToString("#,####");
            respone.freePreConsAmortYr5 = Math.Round(Convert.ToDecimal(respone.freePreConsAmortYr5, cultures), 0).ToString("#,####");

            respone.freeAcqCostAmort1 = Math.Round(Convert.ToDecimal(respone.freeAcqCostAmort1, cultures), 0).ToString("#,####");
            respone.freeAcqCostAmort2 = Math.Round(Convert.ToDecimal(respone.freeAcqCostAmort2, cultures), 0).ToString("#,####");
            respone.freeAcqCostAmort3 = Math.Round(Convert.ToDecimal(respone.freeAcqCostAmort3, cultures), 0).ToString("#,####");
            respone.freeAcqCostAmort4 = Math.Round(Convert.ToDecimal(respone.freeAcqCostAmort4, cultures), 0).ToString("#,####");
            respone.freeAcqCostAmort5 = Math.Round(Convert.ToDecimal(respone.freeAcqCostAmort5, cultures), 0).ToString("#,####");

            respone.freeGoddwillAmortYr1 = Math.Round(Convert.ToDecimal(respone.freeGoddwillAmortYr1, cultures), 0).ToString("#,####");
            respone.freeGoddwillAmortYr2 = Math.Round(Convert.ToDecimal(respone.freeGoddwillAmortYr2, cultures), 0).ToString("#,####");
            respone.freeGoddwillAmortYr3 = Math.Round(Convert.ToDecimal(respone.freeGoddwillAmortYr3, cultures), 0).ToString("#,####");
            respone.freeGoddwillAmortYr4 = Math.Round(Convert.ToDecimal(respone.freeGoddwillAmortYr4, cultures), 0).ToString("#,####");
            respone.freeGoddwillAmortYr5 = Math.Round(Convert.ToDecimal(respone.freeGoddwillAmortYr5, cultures), 0).ToString("#,####");

            respone.freeWCchangeYr1 = Math.Round(Convert.ToDecimal(respone.freeWCchangeYr1, cultures), 0).ToString("#,####");
            respone.freeWCchangeYr2 = Math.Round(Convert.ToDecimal(respone.freeWCchangeYr2, cultures), 0).ToString("#,####");
            respone.freeWCchangeYr3 = Math.Round(Convert.ToDecimal(respone.freeWCchangeYr3, cultures), 0).ToString("#,####");
            respone.freeWCchangeYr4 = Math.Round(Convert.ToDecimal(respone.freeWCchangeYr4, cultures), 0).ToString("#,####");
            respone.freeWCchangeYr5 = Math.Round(Convert.ToDecimal(respone.freeWCchangeYr5, cultures), 0).ToString("#,####");

            respone.freeRevPaydownWCYr1 = Math.Round(Convert.ToDecimal(respone.freeRevPaydownWCYr1, cultures), 0).ToString("#,####");
            respone.freeRevPaydownWCYr2 = Math.Round(Convert.ToDecimal(respone.freeRevPaydownWCYr2, cultures), 0).ToString("#,####");
            respone.freeRevPaydownWCYr3 = Math.Round(Convert.ToDecimal(respone.freeRevPaydownWCYr3, cultures), 0).ToString("#,####");
            respone.freeRevPaydownWCYr4 = Math.Round(Convert.ToDecimal(respone.freeRevPaydownWCYr4, cultures), 0).ToString("#,####");
            respone.freeRevPaydownWCYr5 = Math.Round(Convert.ToDecimal(respone.freeRevPaydownWCYr5, cultures), 0).ToString("#,####");

            respone.freeTermLoanPayYr1 = Math.Round(Convert.ToDecimal(respone.freeTermLoanPayYr1, cultures), 0).ToString("#,####");
            respone.freeTermLoanPayYr2 = Math.Round(Convert.ToDecimal(respone.freeTermLoanPayYr2, cultures), 0).ToString("#,####");
            respone.freeTermLoanPayYr3 = Math.Round(Convert.ToDecimal(respone.freeTermLoanPayYr3, cultures), 0).ToString("#,####");
            respone.freeTermLoanPayYr4 = Math.Round(Convert.ToDecimal(respone.freeTermLoanPayYr4, cultures), 0).ToString("#,####");
            respone.freeTermLoanPayYr5 = Math.Round(Convert.ToDecimal(respone.freeTermLoanPayYr5, cultures), 0).ToString("#,####");

            respone.freeOverAdvLoanPayYr1 = Math.Round(Convert.ToDecimal(respone.freeOverAdvLoanPayYr1, cultures), 0).ToString("#,####");
            respone.freeOverAdvLoanPayYr2 = Math.Round(Convert.ToDecimal(respone.freeOverAdvLoanPayYr2, cultures), 0).ToString("#,####");
            respone.freeOverAdvLoanPayYr3 = Math.Round(Convert.ToDecimal(respone.freeOverAdvLoanPayYr3, cultures), 0).ToString("#,####");
            respone.freeOverAdvLoanPayYr4 = Math.Round(Convert.ToDecimal(respone.freeOverAdvLoanPayYr4, cultures), 0).ToString("#,####");
            respone.freeOverAdvLoanPayYr5 = Math.Round(Convert.ToDecimal(respone.freeOverAdvLoanPayYr5, cultures), 0).ToString("#,####");

            respone.freeMezzFinanPayYr1 = Math.Round(Convert.ToDecimal(respone.freeMezzFinanPayYr1, cultures), 0).ToString("#,####");
            respone.freeMezzFinanPayYr2 = Math.Round(Convert.ToDecimal(respone.freeMezzFinanPayYr2, cultures), 0).ToString("#,####");
            respone.freeMezzFinanPayYr3 = Math.Round(Convert.ToDecimal(respone.freeMezzFinanPayYr3, cultures), 0).ToString("#,####");
            respone.freeMezzFinanPayYr4 = Math.Round(Convert.ToDecimal(respone.freeMezzFinanPayYr4, cultures), 0).ToString("#,####");
           // respone.freeMezzFinanPayYr5 = Math.Round(Convert.ToDecimal(respone.freeMezzFinanPayYr5, cultures), 0).ToString("#,####");

            respone.freeGapNotePayYr1 = Math.Round(Convert.ToDecimal(respone.freeGapNotePayYr1, cultures), 0).ToString("#,####");
            respone.freeGapNotePayYr2 = Math.Round(Convert.ToDecimal(respone.freeGapNotePayYr2, cultures), 0).ToString("#,####");
            respone.freeGapNotePayYr3 = Math.Round(Convert.ToDecimal(respone.freeGapNotePayYr3, cultures), 0).ToString("#,####");
            respone.freeGapNotePayYr4 = Math.Round(Convert.ToDecimal(respone.freeGapNotePayYr4, cultures), 0).ToString("#,####");
            respone.freeGapNotePayYr5 = Math.Round(Convert.ToDecimal(respone.freeGapNotePayYr5, cultures), 0).ToString("#,####");

            respone.freeGapBalloonNotePayYr1 = Math.Round(Convert.ToDecimal(respone.freeGapBalloonNotePayYr1, cultures), 0).ToString("#,####");
            respone.freeGapBalloonNotePayYr2 = Math.Round(Convert.ToDecimal(respone.freeGapBalloonNotePayYr2, cultures), 0).ToString("#,####");
            respone.freeGapBalloonNotePayYr3 = Math.Round(Convert.ToDecimal(respone.freeGapBalloonNotePayYr3, cultures), 0).ToString("#,####");
            respone.freeGapBalloonNotePayYr4 = Math.Round(Convert.ToDecimal(respone.freeGapBalloonNotePayYr4, cultures), 0).ToString("#,####");
            respone.freeGapBalloonNotePayYr5 = Math.Round(Convert.ToDecimal(respone.freeGapBalloonNotePayYr5, cultures), 0).ToString("#,####");

            respone.freeRemNonCompPayYr1 = Math.Round(Convert.ToDecimal(respone.freeRemNonCompPayYr1, cultures), 0).ToString("#,####");
            respone.freeRemNonCompPayYr2 = Math.Round(Convert.ToDecimal(respone.freeRemNonCompPayYr2, cultures), 0).ToString("#,####");
            respone.freeRemNonCompPayYr3 = Math.Round(Convert.ToDecimal(respone.freeRemNonCompPayYr3, cultures), 0).ToString("#,####");
            respone.freeRemNonCompPayYr4 = Math.Round(Convert.ToDecimal(respone.freeRemNonCompPayYr4, cultures), 0).ToString("#,####");
            respone.freeRemNonCompPayYr5 = Math.Round(Convert.ToDecimal(respone.freeRemNonCompPayYr5, cultures), 0).ToString("#,####");

            respone.freeRemPersGoodwillPayYr1 = Math.Round(Convert.ToDecimal(respone.freeRemPersGoodwillPayYr1, cultures), 0).ToString("#,####");
            respone.freeRemPersGoodwillPayYr2 = Math.Round(Convert.ToDecimal(respone.freeRemPersGoodwillPayYr2, cultures), 0).ToString("#,####");
            respone.freeRemPersGoodwillPayYr3 = Math.Round(Convert.ToDecimal(respone.freeRemPersGoodwillPayYr3, cultures), 0).ToString("#,####");
            respone.freeRemPersGoodwillPayYr4 = Math.Round(Convert.ToDecimal(respone.freeRemPersGoodwillPayYr4, cultures), 0).ToString("#,####");
            respone.freeRemPersGoodwillPayYr5 = Math.Round(Convert.ToDecimal(respone.freeRemPersGoodwillPayYr5, cultures), 0).ToString("#,####");

            respone.freeCapExpYr1 = Math.Round(Convert.ToDecimal(respone.freeCapExpYr1, cultures), 0).ToString("#,####");
            respone.freeCapExpYr2 = Math.Round(Convert.ToDecimal(respone.freeCapExpYr2, cultures), 0).ToString("#,####");
            respone.freeCapExpYr3 = Math.Round(Convert.ToDecimal(respone.freeCapExpYr3, cultures), 0).ToString("#,####");
            respone.freeCapExpYr4 = Math.Round(Convert.ToDecimal(respone.freeCapExpYr4, cultures), 0).ToString("#,####");
            respone.freeCapExpYr5 = Math.Round(Convert.ToDecimal(respone.freeCapExpYr5, cultures), 0).ToString("#,####");

            respone.freeCapExpBorrowYr1 = Math.Round(Convert.ToDecimal(respone.freeCapExpBorrowYr1, cultures), 0).ToString("#,####");
            respone.freeCapExpBorrowYr2 = Math.Round(Convert.ToDecimal(respone.freeCapExpBorrowYr2, cultures), 0).ToString("#,####");
            respone.freeCapExpBorrowYr3 = Math.Round(Convert.ToDecimal(respone.freeCapExpBorrowYr3, cultures), 0).ToString("#,####");
            respone.freeCapExpBorrowYr4 = Math.Round(Convert.ToDecimal(respone.freeCapExpBorrowYr4, cultures), 0).ToString("#,####");
            respone.freeCapExpBorrowYr5 = Math.Round(Convert.ToDecimal(respone.freeCapExpBorrowYr5, cultures), 0).ToString("#,####");

            respone.freeCapExpPayYr1 = Math.Round(Convert.ToDecimal(respone.freeCapExpPayYr1, cultures), 0).ToString("#,####");
            respone.freeCapExpPayYr2 = Math.Round(Convert.ToDecimal(respone.freeCapExpPayYr2, cultures), 0).ToString("#,####");
            respone.freeCapExpPayYr3 = Math.Round(Convert.ToDecimal(respone.freeCapExpPayYr3, cultures), 0).ToString("#,####");
            respone.freeCapExpPayYr4 = Math.Round(Convert.ToDecimal(respone.freeCapExpPayYr4, cultures), 0).ToString("#,####");
            respone.freeCapExpPayYr5 = Math.Round(Convert.ToDecimal(respone.freeCapExpPayYr5, cultures), 0).ToString("#,####");

            respone.freeEarnOutPayPriceAdjYr1 = Math.Round(Convert.ToDecimal(respone.freeEarnOutPayPriceAdjYr1, cultures), 0).ToString("#,####");
            respone.freeEarnOutPayPriceAdjYr2 = Math.Round(Convert.ToDecimal(respone.freeEarnOutPayPriceAdjYr2, cultures), 0).ToString("#,####");
            respone.freeEarnOutPayPriceAdjYr3 = Math.Round(Convert.ToDecimal(respone.freeEarnOutPayPriceAdjYr3, cultures), 0).ToString("#,####");
            respone.freeEarnOutPayPriceAdjYr4 = Math.Round(Convert.ToDecimal(respone.freeEarnOutPayPriceAdjYr4, cultures), 0).ToString("#,####");
            respone.freeEarnOutPayPriceAdjYr5 = Math.Round(Convert.ToDecimal(respone.freeEarnOutPayPriceAdjYr5, cultures), 0).ToString("#,####");

            respone.freeDisShareholdertaxyr1 = Math.Round(Convert.ToDecimal(respone.freeDisShareholdertaxyr1, cultures), 0).ToString("#,####");
            respone.freeDisShareholdertaxyr2 = Math.Round(Convert.ToDecimal(respone.freeDisShareholdertaxyr2, cultures), 0).ToString("#,####");
            respone.freeDisShareholdertaxyr3 = Math.Round(Convert.ToDecimal(respone.freeDisShareholdertaxyr3, cultures), 0).ToString("#,####");
            respone.freeDisShareholdertaxyr4 = Math.Round(Convert.ToDecimal(respone.freeDisShareholdertaxyr4, cultures), 0).ToString("#,####");
            respone.freeDisShareholdertaxyr5 = Math.Round(Convert.ToDecimal(respone.freeDisShareholdertaxyr5, cultures), 0).ToString("#,####");

            respone.freeOpeCashFlowBusiYr1 = Math.Round(Convert.ToDecimal(respone.freeOpeCashFlowBusiYr1, cultures), 0).ToString("#,####");
            respone.freeOpeCashFlowBusiYr2 = Math.Round(Convert.ToDecimal(respone.freeOpeCashFlowBusiYr2, cultures), 0).ToString("#,####");
            respone.freeOpeCashFlowBusiYr3 = Math.Round(Convert.ToDecimal(respone.freeOpeCashFlowBusiYr3, cultures), 0).ToString("#,####");
            respone.freeOpeCashFlowBusiYr4 = Math.Round(Convert.ToDecimal(respone.freeOpeCashFlowBusiYr4, cultures), 0).ToString("#,####");
            respone.freeOpeCashFlowBusiYr5 = Math.Round(Convert.ToDecimal(respone.freeOpeCashFlowBusiYr5, cultures), 0).ToString("#,####");

            respone.freeOpeCashFlowRealEstateYr1 = Math.Round(Convert.ToDecimal(respone.freeOpeCashFlowRealEstateYr1, cultures), 0).ToString("#,####");
            respone.freeOpeCashFlowRealEstateYr2 = Math.Round(Convert.ToDecimal(respone.freeOpeCashFlowRealEstateYr2, cultures), 0).ToString("#,####");
            respone.freeOpeCashFlowRealEstateYr3 = Math.Round(Convert.ToDecimal(respone.freeOpeCashFlowRealEstateYr3, cultures), 0).ToString("#,####");
            respone.freeOpeCashFlowRealEstateYr4 = Math.Round(Convert.ToDecimal(respone.freeOpeCashFlowRealEstateYr4, cultures), 0).ToString("#,####");
            respone.freeOpeCashFlowRealEstateYr5 = Math.Round(Convert.ToDecimal(respone.freeOpeCashFlowRealEstateYr5, cultures), 0).ToString("#,####");

            respone.freeOperCashFlowTotalYr1 = Math.Round(Convert.ToDecimal(respone.freeOperCashFlowTotalYr1, cultures), 0).ToString("#,####");
            respone.freeOperCashFlowTotalYr2 = Math.Round(Convert.ToDecimal(respone.freeOperCashFlowTotalYr2, cultures), 0).ToString("#,####");
            respone.freeOperCashFlowTotalYr3 = Math.Round(Convert.ToDecimal(respone.freeOperCashFlowTotalYr3, cultures), 0).ToString("#,####");
            respone.freeOperCashFlowTotalYr4 = Math.Round(Convert.ToDecimal(respone.freeOperCashFlowTotalYr4, cultures), 0).ToString("#,####");
            respone.freeOperCashFlowTotalYr5 = Math.Round(Convert.ToDecimal(respone.freeOperCashFlowTotalYr5, cultures), 0).ToString("#,####");

            respone.freebeginCashBalYr1 = Math.Round(Convert.ToDecimal(respone.freebeginCashBalYr1, cultures), 0).ToString("#,####");
            respone.freebeginCashBalYr2 = Math.Round(Convert.ToDecimal(respone.freebeginCashBalYr2, cultures), 0).ToString("#,####");
            respone.freebeginCashBalYr3 = Math.Round(Convert.ToDecimal(respone.freebeginCashBalYr3, cultures), 0).ToString("#,####");
            respone.freebeginCashBalYr4 = Math.Round(Convert.ToDecimal(respone.freebeginCashBalYr4, cultures), 0).ToString("#,####");
            respone.freebeginCashBalYr5 = Math.Round(Convert.ToDecimal(respone.freebeginCashBalYr5, cultures), 0).ToString("#,####");

            respone.freeOperCashFlowYr1 = Math.Round(Convert.ToDecimal(respone.freeOperCashFlowYr1, cultures), 0).ToString("#,####");
            respone.freeOperCashFlowYr2 = Math.Round(Convert.ToDecimal(respone.freeOperCashFlowYr2, cultures), 0).ToString("#,####");
            respone.freeOperCashFlowYr3 = Math.Round(Convert.ToDecimal(respone.freeOperCashFlowYr3, cultures), 0).ToString("#,####");
            respone.freeOperCashFlowYr4 = Math.Round(Convert.ToDecimal(respone.freeOperCashFlowYr4, cultures), 0).ToString("#,####");
            respone.freeOperCashFlowYr5 = Math.Round(Convert.ToDecimal(respone.freeOperCashFlowYr5, cultures), 0).ToString("#,####");

            respone.freeOpeCashReqYr1 = Math.Round(Convert.ToDecimal(respone.freeOpeCashReqYr1, cultures), 0).ToString("#,####");
            respone.freeOpeCashReqYr2 = Math.Round(Convert.ToDecimal(respone.freeOpeCashReqYr2, cultures), 0).ToString("#,####");
            respone.freeOpeCashReqYr3 = Math.Round(Convert.ToDecimal(respone.freeOpeCashReqYr3, cultures), 0).ToString("#,####");
            respone.freeOpeCashReqYr4 = Math.Round(Convert.ToDecimal(respone.freeOpeCashReqYr4, cultures), 0).ToString("#,####");
            respone.freeOpeCashReqYr5 = Math.Round(Convert.ToDecimal(respone.freeOpeCashReqYr5, cultures), 0).ToString("#,####");

            respone.freeCashRevEndYr1 = Math.Round(Convert.ToDecimal(respone.freeCashRevEndYr1, cultures), 0).ToString("#,####");
            respone.freeCashRevEndYr2 = Math.Round(Convert.ToDecimal(respone.freeCashRevEndYr2, cultures), 0).ToString("#,####");
            respone.freeCashRevEndYr3 = Math.Round(Convert.ToDecimal(respone.freeCashRevEndYr3, cultures), 0).ToString("#,####");
            respone.freeCashRevEndYr4 = Math.Round(Convert.ToDecimal(respone.freeCashRevEndYr4, cultures), 0).ToString("#,####");
            respone.freeCashRevEndYr5 = Math.Round(Convert.ToDecimal(respone.freeCashRevEndYr5, cultures), 0).ToString("#,####");

            respone.freeb4BorrowingYr1 = Math.Round(Convert.ToDecimal(respone.freeb4BorrowingYr1, cultures), 0).ToString("#,####");
            respone.freeb4BorrowingYr2 = Math.Round(Convert.ToDecimal(respone.freeb4BorrowingYr2, cultures), 0).ToString("#,####");
            respone.freeb4BorrowingYr3 = Math.Round(Convert.ToDecimal(respone.freeb4BorrowingYr3, cultures), 0).ToString("#,####");
            respone.freeb4BorrowingYr4 = Math.Round(Convert.ToDecimal(respone.freeb4BorrowingYr4, cultures), 0).ToString("#,####");
            respone.freeb4BorrowingYr5 = Math.Round(Convert.ToDecimal(respone.freeb4BorrowingYr5, cultures), 0).ToString("#,####");

            respone.freeAvailCreditLineyr1 = Math.Round(Convert.ToDecimal(respone.freeAvailCreditLineyr1, cultures), 0).ToString("#,####");
            respone.freeAvailCreditLineyr2 = Math.Round(Convert.ToDecimal(respone.freeAvailCreditLineyr2, cultures), 0).ToString("#,####");
            respone.freeAvailCreditLineyr3 = Math.Round(Convert.ToDecimal(respone.freeAvailCreditLineyr3, cultures), 0).ToString("#,####");
            respone.freeAvailCreditLineyr4 = Math.Round(Convert.ToDecimal(respone.freeAvailCreditLineyr4, cultures), 0).ToString("#,####");
            respone.freeAvailCreditLineyr5 = Math.Round(Convert.ToDecimal(respone.freeAvailCreditLineyr5, cultures), 0).ToString("#,####");

            respone.freeAddRevolverYr1 = Math.Round(Convert.ToDecimal(respone.freeAddRevolverYr1, cultures), 0).ToString("#,####");
            respone.freeAddRevolverYr2 = Math.Round(Convert.ToDecimal(respone.freeAddRevolverYr2, cultures), 0).ToString("#,####");
            respone.freeAddRevolverYr3 = Math.Round(Convert.ToDecimal(respone.freeAddRevolverYr3, cultures), 0).ToString("#,####");
            respone.freeAddRevolverYr4 = Math.Round(Convert.ToDecimal(respone.freeAddRevolverYr4, cultures), 0).ToString("#,####");
            respone.freeAddRevolverYr5 = Math.Round(Convert.ToDecimal(respone.freeAddRevolverYr5, cultures), 0).ToString("#,####");

            respone.freeBVXCashFlowYr1 = Math.Round(Convert.ToDecimal(respone.freeBVXCashFlowYr1, cultures), 0).ToString("#,####");
            respone.freeBVXCashFlowYr2 = Math.Round(Convert.ToDecimal(respone.freeBVXCashFlowYr2, cultures), 0).ToString("#,####");
            respone.freeBVXCashFlowYr3 = Math.Round(Convert.ToDecimal(respone.freeBVXCashFlowYr3, cultures), 0).ToString("#,####");
            respone.freeBVXCashFlowYr4 = Math.Round(Convert.ToDecimal(respone.freeBVXCashFlowYr4, cultures), 0).ToString("#,####");
            respone.freeBVXCashFlowYr5 = Math.Round(Convert.ToDecimal(respone.freeBVXCashFlowYr5, cultures), 0).ToString("#,####");

            respone.freeAddCapContributeYr1 = Math.Round(Convert.ToDecimal(respone.freeAddCapContributeYr1, cultures), 0).ToString("#,####");
            respone.freeAddCapContributeYr2 = Math.Round(Convert.ToDecimal(respone.freeAddCapContributeYr2, cultures), 0).ToString("#,####");
            respone.freeAddCapContributeYr3 = Math.Round(Convert.ToDecimal(respone.freeAddCapContributeYr3, cultures), 0).ToString("#,####");
            respone.freeAddCapContributeYr4 = Math.Round(Convert.ToDecimal(respone.freeAddCapContributeYr4, cultures), 0).ToString("#,####");
            respone.freeAddCapContributeYr5 = Math.Round(Convert.ToDecimal(respone.freeAddCapContributeYr5, cultures), 0).ToString("#,####");

            respone.freeDividendDistrRegYr1 = Math.Round(Convert.ToDecimal(respone.freeDividendDistrRegYr1, cultures), 0).ToString("#,####");
            respone.freeDividendDistrRegYr2 = Math.Round(Convert.ToDecimal(respone.freeDividendDistrRegYr2, cultures), 0).ToString("#,####");
            respone.freeDividendDistrRegYr3 = Math.Round(Convert.ToDecimal(respone.freeDividendDistrRegYr3, cultures), 0).ToString("#,####");
            respone.freeDividendDistrRegYr4 = Math.Round(Convert.ToDecimal(respone.freeDividendDistrRegYr4, cultures), 0).ToString("#,####");
            respone.freeDividendDistrRegYr5 = Math.Round(Convert.ToDecimal(respone.freeDividendDistrRegYr5, cultures), 0).ToString("#,####");

            respone.freeAddOverAdvLoanPayYr1 = Math.Round(Convert.ToDecimal(respone.freeAddOverAdvLoanPayYr1, cultures), 0).ToString("#,####");
            respone.freeAddOverAdvLoanPayYr2 = Math.Round(Convert.ToDecimal(respone.freeAddOverAdvLoanPayYr2, cultures), 0).ToString("#,####");
            respone.freeAddOverAdvLoanPayYr3 = Math.Round(Convert.ToDecimal(respone.freeAddOverAdvLoanPayYr3, cultures), 0).ToString("#,####");
            respone.freeAddOverAdvLoanPayYr4 = Math.Round(Convert.ToDecimal(respone.freeAddOverAdvLoanPayYr4, cultures), 0).ToString("#,####");
            respone.freeAddOverAdvLoanPayYr5 = Math.Round(Convert.ToDecimal(respone.freeAddOverAdvLoanPayYr5, cultures), 0).ToString("#,####");

            respone.freeAddRevolverPayYr1 = Math.Round(Convert.ToDecimal(respone.freeAddRevolverPayYr1, cultures), 0).ToString("#,####");
            respone.freeAddRevolverPayYr2 = Math.Round(Convert.ToDecimal(respone.freeAddRevolverPayYr2, cultures), 0).ToString("#,####");
            respone.freeAddRevolverPayYr3 = Math.Round(Convert.ToDecimal(respone.freeAddRevolverPayYr3, cultures), 0).ToString("#,####");
            respone.freeAddRevolverPayYr4 = Math.Round(Convert.ToDecimal(respone.freeAddRevolverPayYr4, cultures), 0).ToString("#,####");
            respone.freeAddRevolverPayYr5 = Math.Round(Convert.ToDecimal(respone.freeAddRevolverPayYr5, cultures), 0).ToString("#,####");

            respone.freeAddTremloanPayYr1 = Math.Round(Convert.ToDecimal(respone.freeAddTremloanPayYr1, cultures), 0).ToString("#,####");
            respone.freeAddTremloanPayYr2 = Math.Round(Convert.ToDecimal(respone.freeAddTremloanPayYr2, cultures), 0).ToString("#,####");
            respone.freeAddTremloanPayYr3 = Math.Round(Convert.ToDecimal(respone.freeAddTremloanPayYr3, cultures), 0).ToString("#,####");
            respone.freeAddTremloanPayYr4 = Math.Round(Convert.ToDecimal(respone.freeAddTremloanPayYr4, cultures), 0).ToString("#,####");
            respone.freeAddTremloanPayYr5 = Math.Round(Convert.ToDecimal(respone.freeAddTremloanPayYr5, cultures), 0).ToString("#,####");

            respone.freeAddNewCapExPayYr1 = Math.Round(Convert.ToDecimal(respone.freeAddNewCapExPayYr1, cultures), 0).ToString("#,####");
            respone.freeAddNewCapExPayYr2 = Math.Round(Convert.ToDecimal(respone.freeAddNewCapExPayYr2, cultures), 0).ToString("#,####");
            respone.freeAddNewCapExPayYr3 = Math.Round(Convert.ToDecimal(respone.freeAddNewCapExPayYr3, cultures), 0).ToString("#,####");
            respone.freeAddNewCapExPayYr4 = Math.Round(Convert.ToDecimal(respone.freeAddNewCapExPayYr4, cultures), 0).ToString("#,####");
            respone.freeAddNewCapExPayYr5 = Math.Round(Convert.ToDecimal(respone.freeAddNewCapExPayYr5, cultures), 0).ToString("#,####");

            respone.freeAddgapNotePayYr1 = Math.Round(Convert.ToDecimal(respone.freeAddgapNotePayYr1, cultures), 0).ToString("#,####");
            respone.freeAddgapNotePayYr2 = Math.Round(Convert.ToDecimal(respone.freeAddgapNotePayYr2, cultures), 0).ToString("#,####");
            respone.freeAddgapNotePayYr3 = Math.Round(Convert.ToDecimal(respone.freeAddgapNotePayYr3, cultures), 0).ToString("#,####");
            respone.freeAddgapNotePayYr4 = Math.Round(Convert.ToDecimal(respone.freeAddgapNotePayYr4, cultures), 0).ToString("#,####");
            respone.freeAddgapNotePayYr5 = Math.Round(Convert.ToDecimal(respone.freeAddgapNotePayYr5, cultures), 0).ToString("#,####");

            respone.freeDiviDistrAddYr1 = Math.Round(Convert.ToDecimal(respone.freeDiviDistrAddYr1, cultures), 0).ToString("#,####");
            respone.freeDiviDistrAddYr2 = Math.Round(Convert.ToDecimal(respone.freeDiviDistrAddYr2, cultures), 0).ToString("#,####");
            respone.freeDiviDistrAddYr3 = Math.Round(Convert.ToDecimal(respone.freeDiviDistrAddYr3, cultures), 0).ToString("#,####");
            respone.freeDiviDistrAddYr4 = Math.Round(Convert.ToDecimal(respone.freeDiviDistrAddYr4, cultures), 0).ToString("#,####");
            respone.freeDiviDistrAddYr5 = Math.Round(Convert.ToDecimal(respone.freeDiviDistrAddYr5, cultures), 0).ToString("#,####");

            respone.freeCapitalFusionYr1 = Math.Round(Convert.ToDecimal(respone.freeCapitalFusionYr1, cultures), 0).ToString("#,####");
            respone.freeCapitalFusionYr2 = Math.Round(Convert.ToDecimal(respone.freeCapitalFusionYr2, cultures), 0).ToString("#,####");
            respone.freeCapitalFusionYr3 = Math.Round(Convert.ToDecimal(respone.freeCapitalFusionYr3, cultures), 0).ToString("#,####");
            respone.freeCapitalFusionYr4 = Math.Round(Convert.ToDecimal(respone.freeCapitalFusionYr4, cultures), 0).ToString("#,####");
            respone.freeCapitalFusionYr5 = Math.Round(Convert.ToDecimal(respone.freeCapitalFusionYr5, cultures), 0).ToString("#,####");

            respone.freeChangeinCashYr1 = Math.Round(Convert.ToDecimal(respone.freeChangeinCashYr1, cultures), 0).ToString("#,####");
            respone.freeChangeinCashYr2 = Math.Round(Convert.ToDecimal(respone.freeChangeinCashYr2, cultures), 0).ToString("#,####");
            respone.freeChangeinCashYr3 = Math.Round(Convert.ToDecimal(respone.freeChangeinCashYr3, cultures), 0).ToString("#,####");
            respone.freeChangeinCashYr4 = Math.Round(Convert.ToDecimal(respone.freeChangeinCashYr4, cultures), 0).ToString("#,####");
            respone.freeChangeinCashYr5 = Math.Round(Convert.ToDecimal(respone.freeChangeinCashYr5, cultures), 0).ToString("#,####");



            // Cash Flow statement



            respone.stmtNetIncomeYr1 = Math.Round(Convert.ToDecimal(respone.stmtNetIncomeYr1, cultures), 0).ToString("#,####");
            respone.stmtNetIncomeYr2 = Math.Round(Convert.ToDecimal(respone.stmtNetIncomeYr2, cultures), 0).ToString("#,####");
            respone.stmtNetIncomeYr3 = Math.Round(Convert.ToDecimal(respone.stmtNetIncomeYr3, cultures), 0).ToString("#,####");
            respone.stmtNetIncomeYr4 = Math.Round(Convert.ToDecimal(respone.stmtNetIncomeYr4, cultures), 0).ToString("#,####");
            respone.stmtNetIncomeYr5 = Math.Round(Convert.ToDecimal(respone.stmtNetIncomeYr5, cultures), 0).ToString("#,####");

            respone.stmtDepriciationYr1 = Math.Round(Convert.ToDecimal(respone.stmtDepriciationYr1, cultures), 0).ToString("#,####");
            respone.stmtDepriciationYr2 = Math.Round(Convert.ToDecimal(respone.stmtDepriciationYr2, cultures), 0).ToString("#,####");
            respone.stmtDepriciationYr3 = Math.Round(Convert.ToDecimal(respone.stmtDepriciationYr3, cultures), 0).ToString("#,####");
            respone.stmtDepriciationYr4 = Math.Round(Convert.ToDecimal(respone.stmtDepriciationYr4, cultures), 0).ToString("#,####");
            respone.stmtDepriciationYr5 = Math.Round(Convert.ToDecimal(respone.stmtDepriciationYr5, cultures), 0).ToString("#,####");

            respone.stmtnonCompAmortyr1 = Math.Round(Convert.ToDecimal(respone.stmtnonCompAmortyr1, cultures), 0).ToString("#,####");
            respone.stmtnonCompAmortyr2 = Math.Round(Convert.ToDecimal(respone.stmtnonCompAmortyr2, cultures), 0).ToString("#,####");
            respone.stmtnonCompAmortyr3 = Math.Round(Convert.ToDecimal(respone.stmtnonCompAmortyr3, cultures), 0).ToString("#,####");
            respone.stmtnonCompAmortyr4 = Math.Round(Convert.ToDecimal(respone.stmtnonCompAmortyr4, cultures), 0).ToString("#,####");
            respone.stmtnonCompAmortyr5 = Math.Round(Convert.ToDecimal(respone.stmtnonCompAmortyr5, cultures), 0).ToString("#,####");

            respone.stmtPerGoodwillAmortYr1 = Math.Round(Convert.ToDecimal(respone.stmtPerGoodwillAmortYr1, cultures), 0).ToString("#,####");
            respone.stmtPerGoodwillAmortYr2 = Math.Round(Convert.ToDecimal(respone.stmtPerGoodwillAmortYr2, cultures), 0).ToString("#,####");
            respone.stmtPerGoodwillAmortYr3 = Math.Round(Convert.ToDecimal(respone.stmtPerGoodwillAmortYr3, cultures), 0).ToString("#,####");
            respone.stmtPerGoodwillAmortYr4 = Math.Round(Convert.ToDecimal(respone.stmtPerGoodwillAmortYr4, cultures), 0).ToString("#,####");
            respone.stmtPerGoodwillAmortYr5 = Math.Round(Convert.ToDecimal(respone.stmtPerGoodwillAmortYr5, cultures), 0).ToString("#,####");

            respone.stmtPreConAmortYr1 = Math.Round(Convert.ToDecimal(respone.stmtPreConAmortYr1, cultures), 0).ToString("#,####");
            respone.stmtPreConAmortYr2 = Math.Round(Convert.ToDecimal(respone.stmtPreConAmortYr2, cultures), 0).ToString("#,####");
            respone.stmtPreConAmortYr3 = Math.Round(Convert.ToDecimal(respone.stmtPreConAmortYr3, cultures), 0).ToString("#,####");
            respone.stmtPreConAmortYr4 = Math.Round(Convert.ToDecimal(respone.stmtPreConAmortYr4, cultures), 0).ToString("#,####");
            respone.stmtPreConAmortYr5 = Math.Round(Convert.ToDecimal(respone.stmtPreConAmortYr5, cultures), 0).ToString("#,####");

            respone.stmtAcqCostAmortYr1 = Math.Round(Convert.ToDecimal(respone.stmtAcqCostAmortYr1, cultures), 0).ToString("#,####");
            respone.stmtAcqCostAmortYr2 = Math.Round(Convert.ToDecimal(respone.stmtAcqCostAmortYr2, cultures), 0).ToString("#,####");
            respone.stmtAcqCostAmortYr3 = Math.Round(Convert.ToDecimal(respone.stmtAcqCostAmortYr3, cultures), 0).ToString("#,####");
            respone.stmtAcqCostAmortYr4 = Math.Round(Convert.ToDecimal(respone.stmtAcqCostAmortYr4, cultures), 0).ToString("#,####");
            respone.stmtAcqCostAmortYr5 = Math.Round(Convert.ToDecimal(respone.stmtAcqCostAmortYr5, cultures), 0).ToString("#,####");

            respone.stmtGoodwillAmortYr1 = Math.Round(Convert.ToDecimal(respone.stmtGoodwillAmortYr1, cultures), 0).ToString("#,####");
            respone.stmtGoodwillAmortYr2 = Math.Round(Convert.ToDecimal(respone.stmtGoodwillAmortYr2, cultures), 0).ToString("#,####");
            respone.stmtGoodwillAmortYr3 = Math.Round(Convert.ToDecimal(respone.stmtGoodwillAmortYr3, cultures), 0).ToString("#,####");
            respone.stmtGoodwillAmortYr4 = Math.Round(Convert.ToDecimal(respone.stmtGoodwillAmortYr4, cultures), 0).ToString("#,####");
            respone.stmtGoodwillAmortYr5 = Math.Round(Convert.ToDecimal(respone.stmtGoodwillAmortYr5, cultures), 0).ToString("#,####");

            respone.stmtWChangeYr1 = Math.Round(Convert.ToDecimal(respone.stmtWChangeYr1, cultures), 0).ToString("#,####");
            respone.stmtWChangeYr2 = Math.Round(Convert.ToDecimal(respone.stmtWChangeYr2, cultures), 0).ToString("#,####");
            respone.stmtWChangeYr3 = Math.Round(Convert.ToDecimal(respone.stmtWChangeYr3, cultures), 0).ToString("#,####");
            respone.stmtWChangeYr4 = Math.Round(Convert.ToDecimal(respone.stmtWChangeYr4, cultures), 0).ToString("#,####");
            respone.stmtWChangeYr5 = Math.Round(Convert.ToDecimal(respone.stmtWChangeYr5, cultures), 0).ToString("#,####");

            respone.stmtOperActivityYr1 = Math.Round(Convert.ToDecimal(respone.stmtOperActivityYr1, cultures), 0).ToString("#,####");
            respone.stmtOperActivityYr2 = Math.Round(Convert.ToDecimal(respone.stmtOperActivityYr2, cultures), 0).ToString("#,####");
            respone.stmtOperActivityYr3 = Math.Round(Convert.ToDecimal(respone.stmtOperActivityYr3, cultures), 0).ToString("#,####");
            respone.stmtOperActivityYr4 = Math.Round(Convert.ToDecimal(respone.stmtOperActivityYr4, cultures), 0).ToString("#,####");
            respone.stmtOperActivityYr5 = Math.Round(Convert.ToDecimal(respone.stmtOperActivityYr5, cultures), 0).ToString("#,####");

            respone.stmtCapitalExpYr1 = Math.Round(Convert.ToDecimal(respone.stmtCapitalExpYr1, cultures), 0).ToString("#,####");
            respone.stmtCapitalExpYr2 = Math.Round(Convert.ToDecimal(respone.stmtCapitalExpYr2, cultures), 0).ToString("#,####");
            respone.stmtCapitalExpYr3 = Math.Round(Convert.ToDecimal(respone.stmtCapitalExpYr3, cultures), 0).ToString("#,####");
            respone.stmtCapitalExpYr4 = Math.Round(Convert.ToDecimal(respone.stmtCapitalExpYr4, cultures), 0).ToString("#,####");
            respone.stmtCapitalExpYr5 = Math.Round(Convert.ToDecimal(respone.stmtCapitalExpYr5, cultures), 0).ToString("#,####");

            respone.stmtPriceAdjYr1 = Math.Round(Convert.ToDecimal(respone.stmtPriceAdjYr1, cultures), 0).ToString("#,####");
            respone.stmtPriceAdjYr2 = Math.Round(Convert.ToDecimal(respone.stmtPriceAdjYr2, cultures), 0).ToString("#,####");
            respone.stmtPriceAdjYr3 = Math.Round(Convert.ToDecimal(respone.stmtPriceAdjYr3, cultures), 0).ToString("#,####");
            respone.stmtPriceAdjYr4 = Math.Round(Convert.ToDecimal(respone.stmtPriceAdjYr4, cultures), 0).ToString("#,####");
            respone.stmtPriceAdjYr5 = Math.Round(Convert.ToDecimal(respone.stmtPriceAdjYr5, cultures), 0).ToString("#,####");

            respone.stmtRemNonCompPayYr1 = Math.Round(Convert.ToDecimal(respone.stmtRemNonCompPayYr1, cultures), 0).ToString("#,####");
            respone.stmtRemNonCompPayYr2 = Math.Round(Convert.ToDecimal(respone.stmtRemNonCompPayYr2, cultures), 0).ToString("#,####");
            respone.stmtRemNonCompPayYr3 = Math.Round(Convert.ToDecimal(respone.stmtRemNonCompPayYr3, cultures), 0).ToString("#,####");
            respone.stmtRemNonCompPayYr4 = Math.Round(Convert.ToDecimal(respone.stmtRemNonCompPayYr4, cultures), 0).ToString("#,####");
            respone.stmtRemNonCompPayYr5 = Math.Round(Convert.ToDecimal(respone.stmtRemNonCompPayYr5, cultures), 0).ToString("#,####");

            respone.stmtRemPerGWPayYr1 = Math.Round(Convert.ToDecimal(respone.stmtRemPerGWPayYr1, cultures), 0).ToString("#,####");
            respone.stmtRemPerGWPayYr2 = Math.Round(Convert.ToDecimal(respone.stmtRemPerGWPayYr2, cultures), 0).ToString("#,####");
            respone.stmtRemPerGWPayYr3 = Math.Round(Convert.ToDecimal(respone.stmtRemPerGWPayYr3, cultures), 0).ToString("#,####");
            respone.stmtRemPerGWPayYr4 = Math.Round(Convert.ToDecimal(respone.stmtRemPerGWPayYr4, cultures), 0).ToString("#,####");
            respone.stmtRemPerGWPayYr5 = Math.Round(Convert.ToDecimal(respone.stmtRemPerGWPayYr5, cultures), 0).ToString("#,####");

            respone.stmtInvActYr1 = Math.Round(Convert.ToDecimal(respone.stmtInvActYr1, cultures), 0).ToString("#,####");
            respone.stmtInvActYr2 = Math.Round(Convert.ToDecimal(respone.stmtInvActYr2, cultures), 0).ToString("#,####");
            respone.stmtInvActYr3 = Math.Round(Convert.ToDecimal(respone.stmtInvActYr3, cultures), 0).ToString("#,####");
            respone.stmtInvActYr4 = Math.Round(Convert.ToDecimal(respone.stmtInvActYr4, cultures), 0).ToString("#,####");
            respone.stmtInvActYr5 = Math.Round(Convert.ToDecimal(respone.stmtAcqCostAmortYr5, cultures), 0).ToString("#,####");

            respone.stmtAddRevolverYr1 = Math.Round(Convert.ToDecimal(respone.stmtAddRevolverYr1, cultures), 0).ToString("#,####");
            respone.stmtAddRevolverYr2 = Math.Round(Convert.ToDecimal(respone.stmtAddRevolverYr2, cultures), 0).ToString("#,####");
            respone.stmtAddRevolverYr3 = Math.Round(Convert.ToDecimal(respone.stmtAddRevolverYr3, cultures), 0).ToString("#,####");
            respone.stmtAddRevolverYr4 = Math.Round(Convert.ToDecimal(respone.stmtAddRevolverYr4, cultures), 0).ToString("#,####");
            respone.stmtAddRevolverYr5 = Math.Round(Convert.ToDecimal(respone.stmtAddRevolverYr5, cultures), 0).ToString("#,####");

            respone.stmtRevPayWCYr1 = Math.Round(Convert.ToDecimal(respone.stmtRevPayWCYr1, cultures), 0).ToString("#,####");
            respone.stmtRevPayWCYr2 = Math.Round(Convert.ToDecimal(respone.stmtRevPayWCYr2, cultures), 0).ToString("#,####");
            respone.stmtRevPayWCYr3 = Math.Round(Convert.ToDecimal(respone.stmtRevPayWCYr3, cultures), 0).ToString("#,####");
            respone.stmtRevPayWCYr4 = Math.Round(Convert.ToDecimal(respone.stmtRevPayWCYr4, cultures), 0).ToString("#,####");
            respone.stmtRevPayWCYr5 = Math.Round(Convert.ToDecimal(respone.stmtRevPayWCYr5, cultures), 0).ToString("#,####");

            respone.stmtAddRevPayYr1 = Math.Round(Convert.ToDecimal(respone.stmtAddRevPayYr1, cultures), 0).ToString("#,####");
            respone.stmtAddRevPayYr2 = Math.Round(Convert.ToDecimal(respone.stmtAddRevPayYr2, cultures), 0).ToString("#,####");
            respone.stmtAddRevPayYr3 = Math.Round(Convert.ToDecimal(respone.stmtAddRevPayYr3, cultures), 0).ToString("#,####");
            respone.stmtAddRevPayYr4 = Math.Round(Convert.ToDecimal(respone.stmtAddRevPayYr4, cultures), 0).ToString("#,####");
            respone.stmtAddRevPayYr5 = Math.Round(Convert.ToDecimal(respone.stmtAddRevPayYr5, cultures), 0).ToString("#,####");

            respone.stmtTermLoanPayYr1 = Math.Round(Convert.ToDecimal(respone.stmtTermLoanPayYr1, cultures), 0).ToString("#,####");
            respone.stmtTermLoanPayYr2 = Math.Round(Convert.ToDecimal(respone.stmtTermLoanPayYr2, cultures), 0).ToString("#,####");
            respone.stmtTermLoanPayYr3 = Math.Round(Convert.ToDecimal(respone.stmtTermLoanPayYr3, cultures), 0).ToString("#,####");
            respone.stmtTermLoanPayYr4 = Math.Round(Convert.ToDecimal(respone.stmtTermLoanPayYr4, cultures), 0).ToString("#,####");
            respone.stmtTermLoanPayYr5 = Math.Round(Convert.ToDecimal(respone.stmtTermLoanPayYr5, cultures), 0).ToString("#,####");

            respone.stmtAddTermLoanpayYr1 = Math.Round(Convert.ToDecimal(respone.stmtAddTermLoanpayYr1, cultures), 0).ToString("#,####");
            respone.stmtAddTermLoanpayYr2 = Math.Round(Convert.ToDecimal(respone.stmtAddTermLoanpayYr2, cultures), 0).ToString("#,####");
            respone.stmtAddTermLoanpayYr3 = Math.Round(Convert.ToDecimal(respone.stmtAddTermLoanpayYr3, cultures), 0).ToString("#,####");
            respone.stmtAddTermLoanpayYr4 = Math.Round(Convert.ToDecimal(respone.stmtAddTermLoanpayYr4, cultures), 0).ToString("#,####");
            respone.stmtAddTermLoanpayYr5 = Math.Round(Convert.ToDecimal(respone.stmtAddTermLoanpayYr5, cultures), 0).ToString("#,####");

            respone.stmtOverAdvLoanPayYr1 = Math.Round(Convert.ToDecimal(respone.stmtOverAdvLoanPayYr1, cultures), 0).ToString("#,####");
            respone.stmtOverAdvLoanPayYr2 = Math.Round(Convert.ToDecimal(respone.stmtOverAdvLoanPayYr2, cultures), 0).ToString("#,####");
            respone.stmtOverAdvLoanPayYr3 = Math.Round(Convert.ToDecimal(respone.stmtOverAdvLoanPayYr3, cultures), 0).ToString("#,####");
            respone.stmtOverAdvLoanPayYr4 = Math.Round(Convert.ToDecimal(respone.stmtOverAdvLoanPayYr4, cultures), 0).ToString("#,####");
            respone.stmtOverAdvLoanPayYr5 = Math.Round(Convert.ToDecimal(respone.stmtOverAdvLoanPayYr5, cultures), 0).ToString("#,####");

            respone.stmtAddOverAdvLoanPayYr1 = Math.Round(Convert.ToDecimal(respone.stmtAddOverAdvLoanPayYr1, cultures), 0).ToString("#,####");
            respone.stmtAddOverAdvLoanPayYr2 = Math.Round(Convert.ToDecimal(respone.stmtAddOverAdvLoanPayYr2, cultures), 0).ToString("#,####");
            respone.stmtAddOverAdvLoanPayYr3 = Math.Round(Convert.ToDecimal(respone.stmtAddOverAdvLoanPayYr3, cultures), 0).ToString("#,####");
            respone.stmtAddOverAdvLoanPayYr4 = Math.Round(Convert.ToDecimal(respone.stmtAddOverAdvLoanPayYr4, cultures), 0).ToString("#,####");
            respone.stmtAddOverAdvLoanPayYr5 = Math.Round(Convert.ToDecimal(respone.stmtAddOverAdvLoanPayYr5, cultures), 0).ToString("#,####");

            *//*respone.stmtMezzFinanPayYr1 = Math.Round(Convert.ToDecimal(respone.stmtMezzFinanPayYr1, cultures), 0).ToString("#,####");
            respone.stmtMezzFinanPayYr2 = Math.Round(Convert.ToDecimal(respone.stmtMezzFinanPayYr2, cultures), 0).ToString("#,####");
            respone.stmtMezzFinanPayYr3 = Math.Round(Convert.ToDecimal(respone.stmtMezzFinanPayYr3, cultures), 0).ToString("#,####");
            respone.stmtMezzFinanPayYr4 = Math.Round(Convert.ToDecimal(respone.stmtMezzFinanPayYr4, cultures), 0).ToString("#,####");
            respone.stmtMezzFinanPayYr5 = Math.Round(Convert.ToDecimal(respone.stmtMezzFinanPayYr5, cultures), 0).ToString("#,####");

            respone.stmtGapNotePayYr1 = Math.Round(Convert.ToDecimal(respone.stmtGapNotePayYr1, cultures), 0).ToString("#,####");
            respone.stmtGapNotePayYr2 = Math.Round(Convert.ToDecimal(respone.stmtGapNotePayYr2, cultures), 0).ToString("#,####");
            respone.stmtGapNotePayYr3 = Math.Round(Convert.ToDecimal(respone.stmtGapNotePayYr3, cultures), 0).ToString("#,####");
            respone.stmtGapNotePayYr4 = Math.Round(Convert.ToDecimal(respone.stmtGapNotePayYr4, cultures), 0).ToString("#,####");
            respone.stmtGapNotePayYr5 = Math.Round(Convert.ToDecimal(respone.stmtGapNotePayYr5, cultures), 0).ToString("#,####");
*//*
            respone.stmtGapBalloonPayYr1 = Math.Round(Convert.ToDecimal(respone.stmtGapBalloonPayYr1, cultures), 0).ToString("#,####");
            respone.stmtGapBalloonPayYr2 = Math.Round(Convert.ToDecimal(respone.stmtGapBalloonPayYr2, cultures), 0).ToString("#,####");
            respone.stmtGapBalloonPayYr3 = Math.Round(Convert.ToDecimal(respone.stmtGapBalloonPayYr3, cultures), 0).ToString("#,####");
            respone.stmtGapBalloonPayYr4 = Math.Round(Convert.ToDecimal(respone.stmtGapBalloonPayYr4, cultures), 0).ToString("#,####");
            respone.stmtGapBalloonPayYr5 = Math.Round(Convert.ToDecimal(respone.stmtGapBalloonPayYr5, cultures), 0).ToString("#,####");

            *//*respone.stmtAddGapNotePayYr1 = Math.Round(Convert.ToDecimal(respone.stmtAddGapNotePayYr1, cultures), 0).ToString("#,####");
            respone.stmtAddGapNotePayYr2 = Math.Round(Convert.ToDecimal(respone.stmtAddGapNotePayYr2, cultures), 0).ToString("#,####");
            respone.stmtAddGapNotePayYr3 = Math.Round(Convert.ToDecimal(respone.stmtAddGapNotePayYr3, cultures), 0).ToString("#,####");
            respone.stmtAddGapNotePayYr4 = Math.Round(Convert.ToDecimal(respone.stmtAddGapNotePayYr4, cultures), 0).ToString("#,####");
            respone.stmtAddGapNotePayYr5 = Math.Round(Convert.ToDecimal(respone.stmtAddGapNotePayYr5, cultures), 0).ToString("#,####");
*//*
            respone.stmtCapExpBorrowYr1 = Math.Round(Convert.ToDecimal(respone.stmtCapExpBorrowYr1, cultures), 0).ToString("#,####");
            respone.stmtCapExpBorrowYr2 = Math.Round(Convert.ToDecimal(respone.stmtCapExpBorrowYr2, cultures), 0).ToString("#,####");
            respone.stmtCapExpBorrowYr3 = Math.Round(Convert.ToDecimal(respone.stmtCapExpBorrowYr3, cultures), 0).ToString("#,####");
            respone.stmtCapExpBorrowYr4 = Math.Round(Convert.ToDecimal(respone.stmtCapExpBorrowYr4, cultures), 0).ToString("#,####");
            respone.stmtCapExpBorrowYr5 = Math.Round(Convert.ToDecimal(respone.stmtCapExpBorrowYr5, cultures), 0).ToString("#,####");

            respone.stmtCapExpPayYr1 = Math.Round(Convert.ToDecimal(respone.stmtCapExpPayYr1, cultures), 0).ToString("#,####");
            respone.stmtCapExpPayYr2 = Math.Round(Convert.ToDecimal(respone.stmtCapExpPayYr2, cultures), 0).ToString("#,####");
            respone.stmtCapExpPayYr3 = Math.Round(Convert.ToDecimal(respone.stmtCapExpPayYr3, cultures), 0).ToString("#,####");
            respone.stmtCapExpPayYr4 = Math.Round(Convert.ToDecimal(respone.stmtCapExpPayYr4, cultures), 0).ToString("#,####");
            respone.stmtCapExpPayYr5 = Math.Round(Convert.ToDecimal(respone.stmtCapExpPayYr5, cultures), 0).ToString("#,####");

            respone.stmtAddNewCapExPayYr1 = Math.Round(Convert.ToDecimal(respone.stmtAddNewCapExPayYr1, cultures), 0).ToString("#,####");
            respone.stmtAddNewCapExPayYr2 = Math.Round(Convert.ToDecimal(respone.stmtAddNewCapExPayYr2, cultures), 0).ToString("#,####");
            respone.stmtAddNewCapExPayYr3 = Math.Round(Convert.ToDecimal(respone.stmtAddNewCapExPayYr3, cultures), 0).ToString("#,####");
            respone.stmtAddNewCapExPayYr4 = Math.Round(Convert.ToDecimal(respone.stmtAddNewCapExPayYr4, cultures), 0).ToString("#,####");
            respone.stmtAddNewCapExPayYr5 = Math.Round(Convert.ToDecimal(respone.stmtAddNewCapExPayYr5, cultures), 0).ToString("#,####");

            respone.stmtAddCapContributeYr1 = Math.Round(Convert.ToDecimal(respone.stmtAddCapContributeYr1, cultures), 0).ToString("#,####");
            respone.stmtAddCapContributeYr2 = Math.Round(Convert.ToDecimal(respone.stmtAddCapContributeYr2, cultures), 0).ToString("#,####");
            respone.stmtAddCapContributeYr3 = Math.Round(Convert.ToDecimal(respone.stmtAddCapContributeYr3, cultures), 0).ToString("#,####");
            respone.stmtAddCapContributeYr4 = Math.Round(Convert.ToDecimal(respone.stmtAddCapContributeYr4, cultures), 0).ToString("#,####");
            respone.stmtAddCapContributeYr5 = Math.Round(Convert.ToDecimal(respone.stmtAddCapContributeYr5, cultures), 0).ToString("#,####");

            respone.stmtDividendDisRegYr1 = Math.Round(Convert.ToDecimal(respone.stmtDividendDisRegYr1, cultures), 0).ToString("#,####");
            respone.stmtDividendDisRegYr2 = Math.Round(Convert.ToDecimal(respone.stmtDividendDisRegYr2, cultures), 0).ToString("#,####");
            respone.stmtDividendDisRegYr3 = Math.Round(Convert.ToDecimal(respone.stmtDividendDisRegYr3, cultures), 0).ToString("#,####");
            respone.stmtDividendDisRegYr4 = Math.Round(Convert.ToDecimal(respone.stmtDividendDisRegYr4, cultures), 0).ToString("#,####");
            respone.stmtDividendDisRegYr5 = Math.Round(Convert.ToDecimal(respone.stmtDividendDisRegYr5, cultures), 0).ToString("#,####");

            respone.stmtDividendDisAddYr1 = Math.Round(Convert.ToDecimal(respone.stmtDividendDisAddYr1, cultures), 0).ToString("#,####");
            respone.stmtDividendDisAddYr2 = Math.Round(Convert.ToDecimal(respone.stmtDividendDisAddYr2, cultures), 0).ToString("#,####");
            respone.stmtDividendDisAddYr3 = Math.Round(Convert.ToDecimal(respone.stmtDividendDisAddYr3, cultures), 0).ToString("#,####");
            respone.stmtDividendDisAddYr4 = Math.Round(Convert.ToDecimal(respone.stmtDividendDisAddYr4, cultures), 0).ToString("#,####");
            respone.stmtDividendDisAddYr5 = Math.Round(Convert.ToDecimal(respone.stmtDividendDisAddYr5, cultures), 0).ToString("#,####");

            respone.stmtDisrShareTaxYr1 = Math.Round(Convert.ToDecimal(respone.stmtDisrShareTaxYr1, cultures), 0).ToString("#,####");
            respone.stmtDisrShareTaxYr2 = Math.Round(Convert.ToDecimal(respone.stmtDisrShareTaxYr2, cultures), 0).ToString("#,####");
            respone.stmtDisrShareTaxYr3 = Math.Round(Convert.ToDecimal(respone.stmtDisrShareTaxYr3, cultures), 0).ToString("#,####");
            respone.stmtDisrShareTaxYr4 = Math.Round(Convert.ToDecimal(respone.stmtDisrShareTaxYr4, cultures), 0).ToString("#,####");
            respone.stmtDisrShareTaxYr5 = Math.Round(Convert.ToDecimal(respone.stmtDisrShareTaxYr5, cultures), 0).ToString("#,####");

            respone.stmtFinanActivitiesYr1 = Math.Round(Convert.ToDecimal(respone.stmtFinanActivitiesYr1, cultures), 0).ToString("#,####");
            respone.stmtFinanActivitiesYr2 = Math.Round(Convert.ToDecimal(respone.stmtFinanActivitiesYr2, cultures), 0).ToString("#,####");
            respone.stmtFinanActivitiesYr3 = Math.Round(Convert.ToDecimal(respone.stmtFinanActivitiesYr3, cultures), 0).ToString("#,####");
            respone.stmtFinanActivitiesYr4 = Math.Round(Convert.ToDecimal(respone.stmtFinanActivitiesYr4, cultures), 0).ToString("#,####");
            respone.stmtFinanActivitiesYr5 = Math.Round(Convert.ToDecimal(respone.stmtFinanActivitiesYr5, cultures), 0).ToString("#,####");

            *//*respone.stmtCashBusinessYr1 = Math.Round(Convert.ToDecimal(respone.stmtCashBusinessYr1, cultures), 0).ToString("#,####");
            respone.stmtCashBusinessYr2 = Math.Round(Convert.ToDecimal(respone.stmtCashBusinessYr2, cultures), 0).ToString("#,####");
            respone.stmtCashBusinessYr3 = Math.Round(Convert.ToDecimal(respone.stmtCashBusinessYr3, cultures), 0).ToString("#,####");
            respone.stmtCashBusinessYr4 = Math.Round(Convert.ToDecimal(respone.stmtCashBusinessYr4, cultures), 0).ToString("#,####");
            respone.stmtCashBusinessYr5 = Math.Round(Convert.ToDecimal(respone.stmtCashBusinessYr5, cultures), 0).ToString("#,####");
*//*
            respone.stmtCashEstateYr1 = Math.Round(Convert.ToDecimal(respone.stmtCashEstateYr1, cultures), 0).ToString("#,####");
            respone.stmtCashEstateYr2 = Math.Round(Convert.ToDecimal(respone.stmtCashEstateYr2, cultures), 0).ToString("#,####");
            respone.stmtCashEstateYr3 = Math.Round(Convert.ToDecimal(respone.stmtCashEstateYr3, cultures), 0).ToString("#,####");
            respone.stmtCashEstateYr4 = Math.Round(Convert.ToDecimal(respone.stmtCashEstateYr4, cultures), 0).ToString("#,####");
            respone.stmtCashEstateYr5 = Math.Round(Convert.ToDecimal(respone.stmtCashEstateYr5, cultures), 0).ToString("#,####");

            *//*respone.stmtTotalinCashYr1 = Math.Round(Convert.ToDecimal(respone.stmtTotalinCashYr1, cultures), 0).ToString("#,####");
            respone.stmtTotalinCashYr2 = Math.Round(Convert.ToDecimal(respone.stmtTotalinCashYr2, cultures), 0).ToString("#,####");
            respone.stmtTotalinCashYr3 = Math.Round(Convert.ToDecimal(respone.stmtTotalinCashYr3, cultures), 0).ToString("#,####");
            respone.stmtTotalinCashYr4 = Math.Round(Convert.ToDecimal(respone.stmtTotalinCashYr4, cultures), 0).ToString("#,####");
            respone.stmtTotalinCashYr5 = Math.Round(Convert.ToDecimal(respone.stmtTotalinCashYr5, cultures), 0).ToString("#,####");
*/
           /* respone.stmtBeginCashBalYr1 = Math.Round(Convert.ToDecimal(respone.stmtBeginCashBalYr1, cultures), 0).ToString("#,####");
            respone.stmtBeginCashBalYr2 = Math.Round(Convert.ToDecimal(respone.stmtBeginCashBalYr2, cultures), 0).ToString("#,####");
            respone.stmtBeginCashBalYr3 = Math.Round(Convert.ToDecimal(respone.stmtBeginCashBalYr3, cultures), 0).ToString("#,####");
            respone.stmtBeginCashBalYr4 = Math.Round(Convert.ToDecimal(respone.stmtBeginCashBalYr4, cultures), 0).ToString("#,####");
            respone.stmtBeginCashBalYr5 = Math.Round(Convert.ToDecimal(respone.stmtBeginCashBalYr5, cultures), 0).ToString("#,####");

            respone.stmtEndCashBalYr1 = Math.Round(Convert.ToDecimal(respone.stmtEndCashBalYr1, cultures), 0).ToString("#,####");
            respone.stmtEndCashBalYr2 = Math.Round(Convert.ToDecimal(respone.stmtEndCashBalYr2, cultures), 0).ToString("#,####");
            respone.stmtEndCashBalYr3 = Math.Round(Convert.ToDecimal(respone.stmtEndCashBalYr3, cultures), 0).ToString("#,####");
            respone.stmtEndCashBalYr4 = Math.Round(Convert.ToDecimal(respone.stmtEndCashBalYr4, cultures), 0).ToString("#,####");
            respone.stmtEndCashBalYr5 = Math.Round(Convert.ToDecimal(respone.stmtEndCashBalYr5, cultures), 0).ToString("#,####");
*//*
            // ROI    


            respone.roiOrigEquInvYr0 = Math.Round(Convert.ToDecimal(respone.roiOrigEquInvYr0, cultures), 2).ToString();

            respone.roiExitMultipleYr5 = Math.Round(Convert.ToDecimal(respone.roiExitMultipleYr5, cultures), 2).ToString();

            respone.roiPlusCashRetainYr5 = Math.Round(Convert.ToDecimal(respone.roiPlusCashRetainYr5, cultures), 2).ToString();

            respone.roiClosingCostExitYr5 = Math.Round(Convert.ToDecimal(respone.roiClosingCostExitYr5, cultures), 2).ToString();

            respone.roiStateYr5 = Math.Round(Convert.ToDecimal(respone.roiStateYr5, cultures), 2).ToString();

            respone.roiLessNonOperLiabYr5 = Math.Round(Convert.ToDecimal(respone.roiLessNonOperLiabYr5, cultures), 2).ToString();

            respone.roiTerminalValYr5 = Math.Round(Convert.ToDecimal(respone.roiTerminalValYr5, cultures), 2).ToString();

            respone.roiShareholderYr1 = Math.Round(Convert.ToDecimal(respone.roiShareholderYr1, cultures), 2).ToString();
            respone.roiShareholderYr2 = Math.Round(Convert.ToDecimal(respone.roiShareholderYr2, cultures), 2).ToString();
            respone.roiShareholderYr3 = Math.Round(Convert.ToDecimal(respone.roiShareholderYr3, cultures), 2).ToString();
            respone.roiShareholderYr4 = Math.Round(Convert.ToDecimal(respone.roiShareholderYr4, cultures), 2).ToString();
            respone.roiShareholderYr5 = Math.Round(Convert.ToDecimal(respone.roiShareholderYr5, cultures), 2).ToString();

            respone.roiIRSYr1 = Math.Round(Convert.ToDecimal(respone.roiIRSYr1, cultures), 2).ToString();
            respone.roiIRSYr2 = Math.Round(Convert.ToDecimal(respone.roiIRSYr2, cultures), 2).ToString();
            respone.roiIRSYr3 = Math.Round(Convert.ToDecimal(respone.roiIRSYr3, cultures), 2).ToString();
            respone.roiIRSYr4 = Math.Round(Convert.ToDecimal(respone.roiIRSYr4, cultures), 2).ToString();
            respone.roiIRSYr5 = Math.Round(Convert.ToDecimal(respone.roiIRSYr5, cultures), 2).ToString();

            respone.roiAddCapContribution1 = Math.Round(Convert.ToDecimal(respone.roiAddCapContribution1, cultures), 2).ToString();
            respone.roiAddCapContribution2 = Math.Round(Convert.ToDecimal(respone.roiAddCapContribution2, cultures), 2).ToString();
            respone.roiAddCapContribution3 = Math.Round(Convert.ToDecimal(respone.roiAddCapContribution3, cultures), 2).ToString();
            respone.roiAddCapContribution4 = Math.Round(Convert.ToDecimal(respone.roiAddCapContribution4, cultures), 2).ToString();
            respone.roiAddCapContribution5 = Math.Round(Convert.ToDecimal(respone.roiAddCapContribution5, cultures), 2).ToString();

            respone.roiDividendDist1 = Math.Round(Convert.ToDecimal(respone.roiDividendDist1, cultures), 2).ToString();
            respone.roiDividendDist2 = Math.Round(Convert.ToDecimal(respone.roiDividendDist2, cultures), 2).ToString();
            respone.roiDividendDist3 = Math.Round(Convert.ToDecimal(respone.roiDividendDist3, cultures), 2).ToString();
            respone.roiDividendDist4 = Math.Round(Convert.ToDecimal(respone.roiDividendDist4, cultures), 2).ToString();
            respone.roiDividendDist5 = Math.Round(Convert.ToDecimal(respone.roiDividendDist5, cultures), 2).ToString();

            respone.roiSdivPreTaxYr1 = Math.Round(Convert.ToDecimal(respone.roiSdivPreTaxYr1, cultures), 2).ToString();
            respone.roiSdivPreTaxYr2 = Math.Round(Convert.ToDecimal(respone.roiSdivPreTaxYr2, cultures), 2).ToString();
            respone.roiSdivPreTaxYr3 = Math.Round(Convert.ToDecimal(respone.roiSdivPreTaxYr3, cultures), 2).ToString();
            respone.roiSdivPreTaxYr4 = Math.Round(Convert.ToDecimal(respone.roiSdivPreTaxYr4, cultures), 2).ToString();
            respone.roiSdivPreTaxYr5 = Math.Round(Convert.ToDecimal(respone.roiSdivPreTaxYr5, cultures), 2).ToString();

            respone.roiUndistributeYr1 = Math.Round(Convert.ToDecimal(respone.roiUndistributeYr1, cultures), 2).ToString();
            respone.roiUndistributeYr2 = Math.Round(Convert.ToDecimal(respone.roiUndistributeYr2, cultures), 2).ToString();
            respone.roiUndistributeYr3 = Math.Round(Convert.ToDecimal(respone.roiUndistributeYr3, cultures), 2).ToString();
            respone.roiUndistributeYr4 = Math.Round(Convert.ToDecimal(respone.roiUndistributeYr4, cultures), 2).ToString();
            respone.roiUndistributeYr5 = Math.Round(Convert.ToDecimal(respone.roiUndistributeYr5, cultures), 2).ToString();

            respone.roiblankYr1 = Math.Round(Convert.ToDecimal(respone.roiblankYr1, cultures), 2).ToString();
            respone.roiblankYr2 = Math.Round(Convert.ToDecimal(respone.roiblankYr2, cultures), 2).ToString();
            respone.roiblankYr3 = Math.Round(Convert.ToDecimal(respone.roiblankYr3, cultures), 2).ToString();
            respone.roiblankYr4 = Math.Round(Convert.ToDecimal(respone.roiblankYr4, cultures), 2).ToString();
            respone.roiblankYr5 = Math.Round(Convert.ToDecimal(respone.roiblankYr5, cultures), 2).ToString();

            respone.roiByerCashFlowyr1 = Math.Round(Convert.ToDecimal(respone.roiByerCashFlowyr1, cultures), 2).ToString();
            respone.roiByerCashFlowyr2 = Math.Round(Convert.ToDecimal(respone.roiByerCashFlowyr2, cultures), 2).ToString();
            respone.roiByerCashFlowyr3 = Math.Round(Convert.ToDecimal(respone.roiByerCashFlowyr3, cultures), 2).ToString();
            respone.roiByerCashFlowyr4 = Math.Round(Convert.ToDecimal(respone.roiByerCashFlowyr4, cultures), 2).ToString();
            respone.roiByerCashFlowyr5 = Math.Round(Convert.ToDecimal(respone.roiByerCashFlowyr5, cultures), 2).ToString();

            respone.roiMezzCashFlowYr0 = Math.Round(Convert.ToDecimal(respone.roiMezzCashFlowYr0, cultures), 2).ToString();
            respone.roiMezzCashFlowYr1 = Math.Round(Convert.ToDecimal(respone.roiMezzCashFlowYr1, cultures), 2).ToString();
            respone.roiMezzCashFlowYr2 = Math.Round(Convert.ToDecimal(respone.roiMezzCashFlowYr2, cultures), 2).ToString();
            respone.roiMezzCashFlowYr3 = Math.Round(Convert.ToDecimal(respone.roiMezzCashFlowYr3, cultures), 2).ToString();
            respone.roiMezzCashFlowYr4 = Math.Round(Convert.ToDecimal(respone.roiMezzCashFlowYr4, cultures), 2).ToString();
            respone.roiMezzCashFlowYr5 = Math.Round(Convert.ToDecimal(respone.roiMezzCashFlowYr5, cultures), 2).ToString();


            respone.roiByerPreTaxROE = Math.Round(Convert.ToDecimal(respone.roiByerPreTaxROE, cultures), 0).ToString("#,####");

            respone.roiMezzFinancing0 = Math.Round(Convert.ToDecimal(respone.roiMezzFinancing0, cultures), 0).ToString("#,####");

            respone.roiIntExpMezzFinanYr0 = Math.Round(Convert.ToDecimal(respone.roiIntExpMezzFinanYr0, cultures), 0).ToString("#,####");
            respone.roiIntExpMezzFinanYr1 = Math.Round(Convert.ToDecimal(respone.roiIntExpMezzFinanYr1, cultures), 0).ToString("#,####");
            respone.roiIntExpMezzFinanYr2 = Math.Round(Convert.ToDecimal(respone.roiIntExpMezzFinanYr2, cultures), 0).ToString("#,####");
            respone.roiIntExpMezzFinanYr3 = Math.Round(Convert.ToDecimal(respone.roiIntExpMezzFinanYr3, cultures), 0).ToString("#,####");
            respone.roiIntExpMezzFinanYr4 = Math.Round(Convert.ToDecimal(respone.roiIntExpMezzFinanYr4, cultures), 0).ToString("#,####");
            respone.roiIntExpMezzFinanYr5 = Math.Round(Convert.ToDecimal(respone.roiIntExpMezzFinanYr5, cultures), 0).ToString("#,####");

            respone.roiMezzPrincipalAmortYr0 = Math.Round(Convert.ToDecimal(respone.roiMezzPrincipalAmortYr0, cultures), 0).ToString("#,####");
            respone.roiMezzPrincipalAmortYr1 = Math.Round(Convert.ToDecimal(respone.roiMezzPrincipalAmortYr1, cultures), 0).ToString("#,####");
            respone.roiMezzPrincipalAmortYr2 = Math.Round(Convert.ToDecimal(respone.roiMezzPrincipalAmortYr2, cultures), 0).ToString("#,####");
            respone.roiMezzPrincipalAmortYr3 = Math.Round(Convert.ToDecimal(respone.roiMezzPrincipalAmortYr3, cultures), 0).ToString("#,####");
            respone.roiMezzPrincipalAmortYr4 = Math.Round(Convert.ToDecimal(respone.roiMezzPrincipalAmortYr4, cultures), 0).ToString("#,####");
            respone.roiMezzPrincipalAmortYr5 = Math.Round(Convert.ToDecimal(respone.roiMezzPrincipalAmortYr5, cultures), 0).ToString("#,####");

            respone.roiMezzRemPrincePayYr0 = Math.Round(Convert.ToDecimal(respone.roiMezzRemPrincePayYr0, cultures), 0).ToString("#,####");
            respone.roiMezzRemPrincePayYr1 = Math.Round(Convert.ToDecimal(respone.roiMezzRemPrincePayYr1, cultures), 0).ToString("#,####");
            respone.roiMezzRemPrincePayYr2 = Math.Round(Convert.ToDecimal(respone.roiMezzRemPrincePayYr2, cultures), 0).ToString("#,####");
            respone.roiMezzRemPrincePayYr3 = Math.Round(Convert.ToDecimal(respone.roiMezzRemPrincePayYr3, cultures), 0).ToString("#,####");
            respone.roiMezzRemPrincePayYr4 = Math.Round(Convert.ToDecimal(respone.roiMezzRemPrincePayYr4, cultures), 0).ToString("#,####");
            respone.roiMezzRemPrincePayYr5 = Math.Round(Convert.ToDecimal(respone.roiMezzRemPrincePayYr5, cultures), 0).ToString("#,####");

            respone.roiMezzSharePreTaxYr0 = Math.Round(Convert.ToDecimal(respone.roiMezzSharePreTaxYr0, cultures), 0).ToString("#,####");
            respone.roiMezzSharePreTaxYr1 = Math.Round(Convert.ToDecimal(respone.roiMezzSharePreTaxYr1, cultures), 0).ToString("#,####");
            respone.roiMezzSharePreTaxYr2 = Math.Round(Convert.ToDecimal(respone.roiMezzSharePreTaxYr2, cultures), 0).ToString("#,####");
            respone.roiMezzSharePreTaxYr3 = Math.Round(Convert.ToDecimal(respone.roiMezzSharePreTaxYr3, cultures), 0).ToString("#,####");
            respone.roiMezzSharePreTaxYr4 = Math.Round(Convert.ToDecimal(respone.roiMezzSharePreTaxYr4, cultures), 0).ToString("#,####");
            respone.roiMezzSharePreTaxYr5 = Math.Round(Convert.ToDecimal(respone.roiMezzSharePreTaxYr5, cultures), 0).ToString("#,####");

            respone.roiMezzaPreTaxCasshFlowYr0 = Math.Round(Convert.ToDecimal(respone.roiMezzaPreTaxCasshFlowYr0, cultures), 0).ToString("#,####");
            respone.roiMezzaPreTaxCasshFlowYr1 = Math.Round(Convert.ToDecimal(respone.roiMezzaPreTaxCasshFlowYr1, cultures), 0).ToString("#,####");
            respone.roiMezzaPreTaxCasshFlowYr2 = Math.Round(Convert.ToDecimal(respone.roiMezzaPreTaxCasshFlowYr2, cultures), 0).ToString("#,####");
            respone.roiMezzaPreTaxCasshFlowYr3 = Math.Round(Convert.ToDecimal(respone.roiMezzaPreTaxCasshFlowYr3, cultures), 0).ToString("#,####");
            respone.roiMezzaPreTaxCasshFlowYr4 = Math.Round(Convert.ToDecimal(respone.roiMezzaPreTaxCasshFlowYr4, cultures), 0).ToString("#,####");
            respone.roiMezzaPreTaxCasshFlowYr5 = Math.Round(Convert.ToDecimal(respone.roiMezzaPreTaxCasshFlowYr5, cultures), 0).ToString("#,####");

            respone.roiActualMezzPreTaxROIYr0 = Math.Round(Convert.ToDecimal(respone.roiActualMezzPreTaxROIYr0, cultures), 0).ToString("#,####");

            respone.roiExpectMezzYr0 = Math.Round(Convert.ToDecimal(respone.roiExpectMezzYr0, cultures), 0).ToString("#,####");
*/

            return JsonContent(respone);
        }

        [HttpPost]
        public ActionResult DownloadExcel(string file)
        {
            string fullPath = Path.Combine(Server.MapPath("~/Excel"), file);
            return File(fullPath, "application/vnd.ms-excel", file);
        }

        private ActionResult JsonContent(object json)
        {
            var camelCaseFormatter = new JsonSerializerSettings();
            camelCaseFormatter.ContractResolver = new CamelCasePropertyNamesContractResolver();
            camelCaseFormatter.NullValueHandling = NullValueHandling.Ignore;
            return Content(JsonConvert.SerializeObject(json, camelCaseFormatter));
        }

        private string GetXml(List<WebSheet> sheets)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(@"<ntvx>");

            foreach (WebSheet sheet in sheets)
            {
                builder.Append(string.Format(@"<sheet name=""{0}"">", sheet.Name));

                foreach (WebCell cell in sheet.Cells)
                {
                    builder.Append(string.Format(@"<cell address= ""{0}"" >{1}</cell>", cell.Key, cell.Value));
                }

                builder.Append("</sheet>");
            }


            builder.Append("</ntvx>");
            return builder.ToString();
        }

        private List<WebSheet> ParseXml(string xml)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xml);

            XmlNode root = xmlDoc.ChildNodes[0];

            List<WebSheet> sheets = new List<WebSheet>();
            WebSheet sheet;
            string sheetName;
            foreach (XmlNode xmlSheet in root.ChildNodes)
            {
                sheetName = xmlSheet.Attributes["name"].Value;
                sheet = sheets.Find(s => s.Name == sheetName);
                if (sheet == null)
                {
                    sheet = new WebSheet() { Name = sheetName, Cells = new List<WebCell>() };
                    sheets.Add(sheet);
                }

                foreach (XmlNode cell in xmlSheet.ChildNodes)
                {
                    sheet.Cells.Add(new WebCell()
                    {
                        Key = cell.Attributes["address"].Value,
                        Value = cell.InnerText
                    });
                }
            }

            return sheets;
        }
    }
}
