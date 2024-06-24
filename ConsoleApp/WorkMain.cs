using System;
using System.IO;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DrillSIM_API.API;
using DrillSIM_API.Packages;

namespace ConsoleApp
{
    class WorkMain
    {
        //Downhole Motor 
        public float CircFlow; //Circulation Flow 1

        // Package: Formation Kick Flow
        public float FRate; // Flow rate 2
        public float FPress; // Formation Pressure 3
        public float FDensity; //Fluid density 4

        // Package: Swab & Surge
        public float SMSpeed; // String Moving Speed 5
        public float BDepth; // Bit Depth 6
        public float CSDepth; //Casing Shoe Depth 7
        public float BSize; // Bit Size 8
        public float MVis; //Mud Viscosity 9

        // Package: Well Control Manager
        public float DPPressure;    //Drill Pipe Pressure 10
        public float CPressure; // Casing Presure 11
        public float FIn; // Flow In 12
        public float FOut; // Flow out 13
        public float MRFlow;        //Mud Return Flow 14
        public float ActiveGL;      //Active gain Loss 15
        public float ATVolume; //Active Tank Volume 16
        public float ROPenetration; //Rate of Penetration 17
        //blic float SPump;  //Strokes Pumped 18
        public float WTBR; //Well TVD Below RKB 18
        public float WMBR; //Well MD Below RKB 19
        public float BTBR; // Bit TVD Below RKB 20
        public float BMBR; // Bit MD Below RKB; 21
        public float MPS1; //22
        public float MPS2; //23
        public float MPS3; //24
        public float AMTD; // Active Mud tank Density; 25
        public float DCL; //Drill Collar length 26
        public float HWDPL; //HWDP Length 27
        public float DPL; //Drill Pipe Length 28
        public float ATMPV; // Active Tank Mud PV; 29
        public float ATMYP; //Active Tank Mud YP; 30
        public float WOBit;         //Weight On Bit 31
        public float HLoad;         //Hook Load 32
        public float WBoPressure; // Well Bottom pressure 33
        public float STP; // Strokes Pumped; 34
        public float WellDepth; //35

        public void Initialise()
        {
            // Do all initialisation here
            FormationFlow.Instance.EnablePackage();
            SwabAndSurge.Instance.EnablePackage();
            WellControlManager.Instance.EnablePackage();
            FrictionLossInAnnulus.Instance.EnablePackage();
            DownholeSloughing.Instance.EnablePackage();
            DownholeDifferentialSticking.Instance.EnablePackage();           
        }

 

        public void Update()
        {
            // Do all updates here
  
            CircFlow = DownholeSloughing.CirculationRate.Get(); //Circulation Rate 1
            //Formation Kick Flow
            FRate = FrictionLossInAnnulus.FlowRate.Get(); //2
            FDensity = FrictionLossInAnnulus.AverageFluidDensity.Get(); //Fluid density 3

            // Package: Swab & Surge
            SMSpeed = SwabAndSurge.StringSpeed.Get(); // String Moving Speed 4
            BDepth = SwabAndSurge.BitDepth.Get(); // Bit Depth 5
            CSDepth = SwabAndSurge.CasingShoeDepth.Get(); //Casing Shoe Depth 6
            BSize =  SwabAndSurge.BitSize.Get(); // Bit Size 7
            MVis = SwabAndSurge.MudViscosity.Get(); //Mud Viscosity 8

            //Well Control Manager
            FPress = WellControlManager.FormationPressure.Get(); //9
            DPPressure = WellControlManager.DrillPipePressure.Get();//10
            CPressure = WellControlManager.CasingPressure.Get();//11
            FIn = WellControlManager.FlowIn.Get();//12
            FOut = WellControlManager.FlowOut.Get();//13
            MRFlow = WellControlManager.GLReturnFlow.Get();//14
            ActiveGL = WellControlManager.GLActiveVolume.Get();//15
            ATVolume = WellControlManager.ActiveTankVolume.Get();//16
            ROPenetration = WellControlManager.ROP.Get();//17
            WOBit = WellControlManager.WeightOnBit.Get();//18
            HLoad = WellControlManager.Hookload.Get();//19
            WBoPressure = WellControlManager.WellBottomPressure.Get();//20
            BTBR = WellControlManager.BitTVD.Get(); // Bit TVD Below RKB 21
            MPS1 = WellControlManager.MudPumps.MP1Speed.Get();//22
            MPS2 = WellControlManager.MudPumps.MP2Speed.Get();//23
            MPS3 = WellControlManager.MudPumps.MP3Speed.Get();//24
            AMTD = WellControlManager.MudDensity.Get(); // Active Mud tank Density; 25
            ATMPV = WellControlManager.MudPV.Get(); // Active Tank Mud PV; 26
            ATMYP = WellControlManager.MudYP.Get(); //Active Tank Mud YP; 27
            STP = WellControlManager.StrokesPumped.Get(); // Strokes Pumped; 28
            WellDepth = WellControlManager.WellTVD.Get(); //29

            Console.WriteLine("wellDepth = {0}, ROP = {1}, Formation Pressure = {2}", WellDepth, ROPenetration, FPress);
            WriteToTextFile(CircFlow.ToString(), FRate.ToString(), FDensity.ToString(), SMSpeed.ToString(), BDepth.ToString(), CSDepth.ToString(), BSize.ToString(), MVis.ToString(), 
                FPress.ToString(), DPPressure.ToString(), CPressure.ToString(), FIn.ToString(), FOut.ToString(), MRFlow.ToString(), ActiveGL.ToString(), ATVolume.ToString(), 
                ROPenetration.ToString(), WOBit.ToString(), HLoad.ToString(), WBoPressure.ToString(), BTBR.ToString(), MPS1.ToString(), MPS2.ToString(), MPS3.ToString(), 
                AMTD.ToString(), ATMPV.ToString(), ATMYP.ToString(), STP.ToString(), WellDepth.ToString());//, RotTorque.ToString(), RotSpeed.ToString()
        }
        public void WriteToTextFile(string CircFlow, string FRate, string FDensity, string SMSpeed, string BDepth, string CSDepth, string BSize, string MVis, string FPress, 
            string DPPressure, string CPressure, string FIn, string FOut, string MRFlow, string ActiveGL, string ATVolume, string ROPenetration, string WOBit, string HLoad, 
            string WBoPressure, string BTBR, string MPS1, string MPS2, string MPS3, string AMTD, string ATMPV, string ATMYP, string STP, string WellDepth) //, string RotTorque, string RotSpeed
        {
            string filePath = "C:\\Projects\\Dipto\\WellControl\\DataFiles\\Data.xlsx"; // Specify the file path
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                int row = worksheet.Dimension.End.Row + 1;

      
                worksheet.Cells[row, 1].Value = CircFlow;
                worksheet.Cells[row, 2].Value = FRate;
                worksheet.Cells[row, 3].Value = FDensity;              
                worksheet.Cells[row, 4].Value = SMSpeed;             
                worksheet.Cells[row, 5].Value = BDepth;
                worksheet.Cells[row, 6].Value = CSDepth;
                worksheet.Cells[row, 7].Value = BSize;
                worksheet.Cells[row, 8].Value = MVis;
                worksheet.Cells[row, 9].Value = FPress;
                worksheet.Cells[row, 10].Value = DPPressure;
                worksheet.Cells[row, 11].Value = CPressure;
                worksheet.Cells[row, 12].Value = FIn;
                worksheet.Cells[row, 13].Value = FOut;
                worksheet.Cells[row, 14].Value = MRFlow;
                worksheet.Cells[row, 15].Value = ActiveGL;
                worksheet.Cells[row, 16].Value = ATVolume;
                worksheet.Cells[row, 17].Value = ROPenetration;
                worksheet.Cells[row, 18].Value = WOBit;
                worksheet.Cells[row, 19].Value = HLoad;
                worksheet.Cells[row, 20].Value = WBoPressure;
                worksheet.Cells[row, 21].Value = BTBR;
                worksheet.Cells[row, 22].Value = MPS1;
                worksheet.Cells[row, 23].Value = MPS2;
                worksheet.Cells[row, 24].Value = MPS3;
                worksheet.Cells[row, 25].Value = AMTD;
                worksheet.Cells[row, 26].Value = ATMPV;
                worksheet.Cells[row, 27].Value = ATMYP;
                worksheet.Cells[row, 28].Value = STP;
                worksheet.Cells[row, 29].Value = WellDepth;       
                package.Save();
            }
        }
    }
    
    
}
