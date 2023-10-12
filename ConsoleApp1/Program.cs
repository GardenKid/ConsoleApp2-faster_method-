using System;

using System.Collections.Generic;

using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Text;
using ConsoleApp1;
using SAP2000v1;
using static ConsoleApp1.classlibrary;

namespace ConsoleApplication1
{

    class Program

    {
        static void Main(string[] args)

        {
            //set the following flag to true to attach to an existing instance of the program
            //otherwise a new instance of the program will be started

            bool AttachToInstance;
            AttachToInstance = true;

            //set the following flag to true to manually specify the path to SAP2000.exe
            //this allows for a connection to a version of SAP2000 other than the latest installation
            //otherwise the latest installed version of SAP2000 will be launched

            bool SpecifyPath;
            SpecifyPath = false;

            //if the above flag is set to true, specify the path to SAP2000 below
            string ProgramPath;
            ProgramPath = @"C:\Program Files\Computers and Structures\SAP2000 24\SAP2000.exe";

            //full path to the model
            //set it to the desired path of your model
            string ModelDirectory = @"E:\虚拟项目\节点cad二次开发\sap2k二次开发\sap2k\111";

            try
            {
                System.IO.Directory.CreateDirectory(ModelDirectory);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Could not create directory: " + ModelDirectory);
            }

            string ModelName = "API_1-001.sdb";

            string ModelPath = ModelDirectory + System.IO.Path.DirectorySeparatorChar + ModelName;

            //dimension the SapObject as cOAPI type

            cOAPI mySapObject = null;

            //Use ret to check if functions return successfully (ret = 0) or fail (ret = nonzero)

            int ret = 0;

            //create API helper object

            cHelper myHelper;

            try
            {
                myHelper = new Helper();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Cannot create an instance of the Helper object");
                return;
            }

            if (AttachToInstance)
            {
                //attach to a running instance of SAP2000
                try
                {
                    //get the active SapObject
                    //The program ID of the API object. Use “CSI.SAP2000.API.SapObject” for SAP2000 and “CSI.CSiBridge.API.SapObject” for CSiBridge.
                    //Attaches to the active running instance of SAP2000 and returns an instance of SapObject if successful, nothing otherwise.
                    mySapObject = myHelper.GetObject("CSI.SAP2000.API.SapObject");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("No running instance of the program found or failed to attach.");
                    return;
                }

            }

            else
            {
                if (SpecifyPath)
                {
                    //'create an instance of the SapObject from the specified path
                    try
                    {
                        //create SapObject

                        mySapObject = myHelper.CreateObject(ProgramPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Cannot start a new instance of the program from " + ProgramPath);

                        return;
                    }

                }

                else
                {
                    //'create an instance of the SapObject from the latest installed SAP2000
                    try
                    {

                        //create SapObject

                        mySapObject = myHelper.CreateObjectProgID("CSI.SAP2000.API.SapObject");

                        //CreateObjectProgID,该函数在注册表中搜索SAP2000的最新安装版本，除非被包含SAP2000.exe完整路径的环境变量覆盖。

                    }
                    catch (Exception ex)
                    {

                        Console.WriteLine("Cannot start a new instance of the program.");

                        return;

                    }

                }

                //start SAP2000 application

                ret = mySapObject.ApplicationStart();

            }



            //create SapModel object

            cSapModel mySapModel = null;


            try
            {
                mySapModel = mySapObject.SapModel;
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
            }    


            //switch to k-in units

            ret = mySapModel.SetPresentUnits(eUnits.N_mm_C);

            // 判断设计结果是否可用，如果不可用就运行分析与设计
            if (mySapModel.DesignSteel.GetResultsAvailable()==false)
            {
                // 保存模型，运行分析与设计
                ret = mySapModel.File.Save(ModelPath);
                ret = mySapModel.Analyze.RunAnalysis();
                // start steel design
                ret = mySapModel.DesignSteel.StartDesign();
            }

            // 读取轴力信息
            //寻找某个对象组的点的所有信息 先创建一个用于存储信息的点结构体列表
            List<classlibrary.PointInfo> PointInfoList_1 = new List<classlibrary.PointInfo>();


            // 依次输入 SapObject 节点对象组名称 节点结构体列表 弦杆对象组名称 腹杆对象组名称
            // 输出 节点结构体列表
            classlibrary.GetPointInfo(mySapObject, "WJ-SX节点", ref PointInfoList_1, "WJ-SX", "Wj-FG");

            //完成弹窗
            Console.WriteLine("完成");
            

            //close SAP2000

            mySapObject.ApplicationExit(false);

        }

    }

}

