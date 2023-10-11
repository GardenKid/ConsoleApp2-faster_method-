using SAP2000v1;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;
using static System.Collections.Specialized.BitVector32;

namespace ConsoleApp1
{
    internal class classlibrary
    {
        public static void GetPointInfo(SAP2000v1.cOAPI SapObject, string GroupName, ref List<classlibrary.PointInfo> PointInfoList_1, string ChordGroupName, string WebGroupName)
        {
            cSapModel SapModel = null;
            //获取SAP2000模型对象
            SapModel = SapObject.SapModel;


            #region 遍历点对象，获取其坐标,以及相连杆件信息
            //获取对象组内的点对象的数量和名称
            int NumberItems = 0;
            int[] ObjectType = null;
            string[] ObjectName = null;
            int ret = SapModel.GroupDef.GetAssignments(GroupName, ref NumberItems, ref ObjectType, ref ObjectName);

            PointInfo temp_PointInfo = new PointInfo();

            for (int i = 0; i < NumberItems; i++)
            {
                Console.WriteLine("开始一个点" + i);
                if (ObjectType[i] == 1)
                {
                    #region 找出对象组里面的所有点
                    //获取点对象的名称
                    string PointName = ObjectName[i];
                    //获取点对象的坐标
                    double pointX = 0;
                    double pointY = 0;
                    double pointZ = 0;
                    ret = SapModel.PointObj.GetCoordCartesian(PointName, ref pointX, ref pointY, ref pointZ);

                    //将点的名称和位置信息写入结构体里面去
                    temp_PointInfo.Name = PointName;
                    temp_PointInfo.pointX = pointX; temp_PointInfo.pointY = pointY; temp_PointInfo.pointZ = pointZ;
                    //  PointLocList_1.Add(temp_PointInfo);
                    // 这一句放到每次循环找点的最后再写，相当于每个点结构体的所有信息都写好了，再将这个点结构体添加到列表中。 
                    #endregion

                    PointInfoList_1.Add(temp_PointInfo);
                }
            }
            #endregion

            // 添加与点相连的弦杆


            //添加与点相连的腹杆

            #region 遍历frame对象，根据其IJ端点名称，将其添加到对应的点结构体list里面去
            // 获取所有杆件的名称
            int numberNames = 0;
            string[] allFrameNames = null;
            //此处如果不定义string[]为null，则下面的语句会报错GetNameList
            SapModel.FrameObj.GetNameList(ref numberNames, ref allFrameNames);

            // 存储与给定点相交的杆件的名称
            List<string> intersectingFrames = new List<string>();


            // 遍历所有杆件，检查是否与给定点相交
            foreach (string frameName in allFrameNames)
            {
                // 提前定义临时FrameInfo结构体，方便后面使用
                List<FrameInfo> temp_FrameInfoList = new List<FrameInfo>();
                FrameInfo temp_FrameInfo = new FrameInfo();

                string point1 = null; string point2 = null;
                SapModel.FrameObj.GetPoints(frameName, ref point1, ref point2);

                #region //判断point1是否在点结构体列表中,point1在列表中，就输出测站2/3，对应编号2的内力。
                bool structExists = PointInfoList_1.Exists(s => s.Name == point1);
                if (structExists)
                {
                    //找出另一个点的坐标放入j点信息中
                    //获取点对象的坐标
                    double X = 0;
                    double Y = 0;
                    double Z = 0;
                    ret = SapModel.PointObj.GetCoordCartesian(point2, ref X, ref Y, ref Z);

                    #region 获取杆件设计内力
                    //获取frame的设计内力对应的组合，然后根据这个组合名字来获取内力
                    int temp_NumberItems = 0;
                    string[] temp_FrameName = new string[1];
                    double[] temp_Ratio = new double[1];
                    int[] temp_RatioType = new int[1];
                    double[] temp_Location = new double[1];
                    string[] temp_ComboName = new string[1];
                    string[] temp_ErrorSummary = new string[1];
                    string[] temp_WarningSummary = new string[1];                    
                    string[] warningSummary = new string[1];

                    ret = SapModel.DesignSteel.GetSummaryResults(frameName, ref temp_NumberItems, ref temp_FrameName, ref temp_Ratio, ref temp_RatioType, ref temp_Location, ref temp_ComboName, ref temp_ErrorSummary, ref temp_WarningSummary, 0);

                    Console.WriteLine(temp_ComboName);
                    //根据这个组合名字来获取内力
                    double[] temp_ItemTypeElm = new double[1];
                    int temp_NumberResults = 0;
                    string[] temp_Obj = new string[1];
                    double[] temp_ObjSta = new double[1];
                    string[] temp_Elm = new string[1];
                    double[] temp_ElmSta = new double[1];
                    string[] temp_LoadCase = temp_ComboName;
                    string[] temp_StepType = new string[1];
                    double[] temp_StepNum = new double[1];   
                    double[] temp_P = new double[1];
                    double[] temp_V2 = new double[1];
                    double[] temp_V3 = new double[1];
                    double[] temp_T = new double[1];
                    double[] temp_M2 = new double[1];
                    double[] temp_M3 = new double[1];

                    //frameforce方法来读取
                    //先设置输出组合
                    SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();
                    ret = SapModel.Results.Setup.SetComboSelectedForOutput(temp_ComboName[0], true);

                    ret = SapModel.Results.FrameForce(frameName, 0, ref temp_NumberResults, ref temp_Obj, ref temp_ObjSta, ref temp_Elm, ref temp_ElmSta, ref temp_LoadCase, ref temp_StepType, ref temp_StepNum, ref temp_P, ref temp_V2, ref temp_V3, ref temp_T, ref temp_M2, ref temp_M3);


                    #endregion


                    //寻找point1对应的点结构体
                    temp_PointInfo = PointInfoList_1.Find(s => s.Name == point1);
                    int index = PointInfoList_1.IndexOf(temp_PointInfo);

                    //将结构体中的列表值赋予给临时列表
                    temp_FrameInfoList = temp_PointInfo.ChordInfoList;

                    //如果temp_PointInfo.WebInfoList;是空的（null），那么需要重新生成下temp_FrameInfoList
                    //因为对null的list进行add，会报错
                    if (temp_FrameInfoList == null)
                    {
                        temp_FrameInfoList = new List<FrameInfo>();
                    }

                    // 修改杆件结构体的成员值
                    temp_FrameInfo.FrameName = frameName;
                    temp_FrameInfo.IName = point1;
                    temp_FrameInfo.JName = point2;
                    temp_FrameInfo.JX = X;
                    temp_FrameInfo.JY = Y;
                    temp_FrameInfo.JZ = Z;
                    temp_FrameInfo.P_max = temp_P.Max();
                    temp_FrameInfo.P_min = temp_P.Min();
                    temp_FrameInfo.P_Cal = (temp_P[2]+ temp_P[1])/2;
                    temp_FrameInfo.ComboName = temp_ComboName[0];
                    temp_FrameInfo.MaxRatio = temp_Ratio[0];

                    // 将杆件结构体加入到对应的点结构体内部的结构体列表中
                    // 分两步，先加入到临时列表中，再将临时列表赋予对应点结构体内部的列表
                    temp_FrameInfoList.Add(temp_FrameInfo);
                    temp_PointInfo.ChordInfoList = temp_FrameInfoList;

                    // 将修改后的结构体重新放回列表中
                    PointInfoList_1[index] = temp_PointInfo;
                }
                #endregion


                #region //判断point2是否在点结构体列表中，point2在列表中，就输出测站5 ，对应编号4的内力。
                structExists = PointInfoList_1.Exists(s => s.Name == point2);
                if (structExists)
                {
                    //找出另一个点的坐标放入j点信息中
                    //获取点对象的坐标
                    double X = 0;
                    double Y = 0;
                    double Z = 0;
                    ret = SapModel.PointObj.GetCoordCartesian(point1, ref X, ref Y, ref Z);

                    #region 获取杆件设计内力
                    //获取frame的设计内力对应的组合，然后根据这个组合名字来获取内力
                    int temp_NumberItems = 0;
                    string[] temp_FrameName = new string[1];
                    double[] temp_Ratio = new double[1];
                    int[] temp_RatioType = new int[1];
                    double[] temp_Location = new double[1];
                    string[] temp_ComboName = new string[1];
                    string[] temp_ErrorSummary = new string[1];
                    string[] temp_WarningSummary = new string[1];
                    string[] warningSummary = new string[1];

                    ret = SapModel.DesignSteel.GetSummaryResults(frameName, ref temp_NumberItems, ref temp_FrameName, ref temp_Ratio, ref temp_RatioType, ref temp_Location, ref temp_ComboName, ref temp_ErrorSummary, ref temp_WarningSummary, 0);

                    Console.WriteLine(temp_ComboName);
                    //根据这个组合名字来获取内力
                    double[] temp_ItemTypeElm = new double[1];
                    int temp_NumberResults = 0;
                    string[] temp_Obj = new string[1];
                    double[] temp_ObjSta = new double[1];
                    string[] temp_Elm = new string[1];
                    double[] temp_ElmSta = new double[1];
                    string[] temp_LoadCase = temp_ComboName;
                    string[] temp_StepType = new string[1];
                    double[] temp_StepNum = new double[1];
                    double[] temp_P = new double[1];
                    double[] temp_V2 = new double[1];
                    double[] temp_V3 = new double[1];
                    double[] temp_T = new double[1];
                    double[] temp_M2 = new double[1];
                    double[] temp_M3 = new double[1];

                    //frameforce方法来读取
                    //先设置输出组合
                    SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();
                    ret = SapModel.Results.Setup.SetComboSelectedForOutput(temp_ComboName[0], true);

                    ret = SapModel.Results.FrameForce(frameName, 0, ref temp_NumberResults, ref temp_Obj, ref temp_ObjSta, ref temp_Elm, ref temp_ElmSta, ref temp_LoadCase, ref temp_StepType, ref temp_StepNum, ref temp_P, ref temp_V2, ref temp_V3, ref temp_T, ref temp_M2, ref temp_M3);


                    #endregion


                    //寻找point2对应的点结构体
                    temp_PointInfo = PointInfoList_1.Find(s => s.Name == point2);
                    int index = PointInfoList_1.IndexOf(temp_PointInfo);

                    //将结构体中的列表值赋予给临时列表
                    temp_FrameInfoList = temp_PointInfo.ChordInfoList;

                    //如果temp_PointInfo.WebInfoList;是空的（null），那么需要重新生成下temp_FrameInfoList
                    //因为对null的list进行add，会报错
                    if (temp_FrameInfoList == null)
                    {
                        temp_FrameInfoList = new List<FrameInfo>();
                    }

                    // 修改杆件结构体的成员值
                    temp_FrameInfo.FrameName = frameName;
                    temp_FrameInfo.IName = point2;
                    temp_FrameInfo.JName = point1;
                    temp_FrameInfo.JX = X;
                    temp_FrameInfo.JY = Y;
                    temp_FrameInfo.JZ = Z;
                    temp_FrameInfo.P_max = temp_P.Max();
                    temp_FrameInfo.P_min = temp_P.Min();
                    temp_FrameInfo.P_Cal = (temp_P[4]+ temp_P[5])/2;
                    temp_FrameInfo.ComboName = temp_ComboName[0];
                    temp_FrameInfo.MaxRatio = temp_Ratio[0];

                    // 将杆件结构体加入到对应的点结构体内部的结构体列表中
                    // 分两步，先加入到临时列表中，再将临时列表赋予对应点结构体内部的列表
                    temp_FrameInfoList.Add(temp_FrameInfo);
                    temp_PointInfo.ChordInfoList = temp_FrameInfoList;

                    // 将修改后的结构体重新放回列表中
                    PointInfoList_1[index] = temp_PointInfo;
                }
                #endregion

                
            }
            #endregion
        }

        //定义方法，输入节点结构体列表，杆件对象组名称，将杆件对象组里面和节点相连的杆件信息加入到对应的结构体里面去
        //如果FrameType是0，则写入到弦杆里面去；如果FrameType是1，则写入到腹杆里面去
        public static void PointListAddFrame(cOAPI SapObject, ref List<classlibrary.PointInfo> PointInfoList_1, string FrameGroup, int FrameType_W_C)
        {

            cSapModel SapModel = null;
            //获取SAP2000模型对象
            SapModel = SapObject.SapModel;

            //变量初始化
            int ret = 0;
            PointInfo temp_PointInfo = new PointInfo();
            int NumberItems = 0;
            int[] ObjectType = null;
            string[] allFrameNames = null;

            // 获取对象组里面所有杆件的名称
            ret = SapModel.GroupDef.GetAssignments(FrameGroup, ref NumberItems, ref ObjectType, ref allFrameNames);


            #region 遍历frame对象，根据其IJ端点名称，将其添加到对应的点结构体list里面去


            // 存储与给定点相交的杆件的名称
            List<string> intersectingFrames = new List<string>();


            // 遍历所有杆件，检查是否与给定点相交
            foreach (string frameName in allFrameNames)
            {
                // 提前定义临时FrameInfo结构体，方便后面使用
                List<FrameInfo> temp_FrameInfoList = new List<FrameInfo>();
                FrameInfo temp_FrameInfo = new FrameInfo();

                string point1 = null; string point2 = null;
                SapModel.FrameObj.GetPoints(frameName, ref point1, ref point2);

                #region //判断point1是否在点结构体列表中,point1在列表中，就输出测站2/3，对应编号2的内力。
                bool structExists = PointInfoList_1.Exists(s => s.Name == point1);
                if (structExists)
                {
                    //找出另一个点的坐标放入j点信息中
                    //获取点对象的坐标
                    double X = 0;
                    double Y = 0;
                    double Z = 0;
                    ret = SapModel.PointObj.GetCoordCartesian(point2, ref X, ref Y, ref Z);

                    #region 获取杆件设计内力
                    //获取frame的设计内力对应的组合，然后根据这个组合名字来获取内力
                    int temp_NumberItems = 0;
                    string[] temp_FrameName = new string[1];
                    double[] temp_Ratio = new double[1];
                    int[] temp_RatioType = new int[1];
                    double[] temp_Location = new double[1];
                    string[] temp_ComboName = new string[1];
                    string[] temp_ErrorSummary = new string[1];
                    string[] temp_WarningSummary = new string[1];
                    string[] warningSummary = new string[1];

                    ret = SapModel.DesignSteel.GetSummaryResults(frameName, ref temp_NumberItems, ref temp_FrameName, ref temp_Ratio, ref temp_RatioType, ref temp_Location, ref temp_ComboName, ref temp_ErrorSummary, ref temp_WarningSummary, 0);

                    Console.WriteLine(temp_ComboName);
                    //根据这个组合名字来获取内力
                    double[] temp_ItemTypeElm = new double[1];
                    int temp_NumberResults = 0;
                    string[] temp_Obj = new string[1];
                    double[] temp_ObjSta = new double[1];
                    string[] temp_Elm = new string[1];
                    double[] temp_ElmSta = new double[1];
                    string[] temp_LoadCase = temp_ComboName;
                    string[] temp_StepType = new string[1];
                    double[] temp_StepNum = new double[1];
                    double[] temp_P = new double[1];
                    double[] temp_V2 = new double[1];
                    double[] temp_V3 = new double[1];
                    double[] temp_T = new double[1];
                    double[] temp_M2 = new double[1];
                    double[] temp_M3 = new double[1];

                    //frameforce方法来读取
                    //先设置输出组合
                    SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();
                    ret = SapModel.Results.Setup.SetComboSelectedForOutput(temp_ComboName[0], true);

                    ret = SapModel.Results.FrameForce(frameName, 0, ref temp_NumberResults, ref temp_Obj, ref temp_ObjSta, ref temp_Elm, ref temp_ElmSta, ref temp_LoadCase, ref temp_StepType, ref temp_StepNum, ref temp_P, ref temp_V2, ref temp_V3, ref temp_T, ref temp_M2, ref temp_M3);


                    #endregion


                    //寻找point1对应的点结构体
                    temp_PointInfo = PointInfoList_1.Find(s => s.Name == point1);
                    int index = PointInfoList_1.IndexOf(temp_PointInfo);

                    //将结构体中的列表值赋予给临时列表
                    //判断杆件是腹杆还是弦杆，分别加入不同的类型中去
                    if (FrameType_W_C == 0)
                    {
                        temp_FrameInfoList = temp_PointInfo.ChordInfoList;
                    }
                    else
                    {
                        temp_FrameInfoList = temp_PointInfo.WebInfoList;
                    }

                    //如果temp_PointInfo.WebInfoList;是空的（null），那么需要重新生成下temp_FrameInfoList
                    //因为对null的list进行add，会报错
                    if (temp_FrameInfoList == null)
                    {
                        temp_FrameInfoList = new List<FrameInfo>();
                    }

                    // 修改杆件结构体的成员值
                    temp_FrameInfo.FrameName = frameName;
                    temp_FrameInfo.IName = point1;
                    temp_FrameInfo.JName = point2;
                    temp_FrameInfo.JX = X;
                    temp_FrameInfo.JY = Y;
                    temp_FrameInfo.JZ = Z;
                    temp_FrameInfo.P_max = temp_P.Max();
                    temp_FrameInfo.P_min = temp_P.Min();
                    temp_FrameInfo.P_Cal = (temp_P[2] + temp_P[1]) / 2;
                    temp_FrameInfo.ComboName = temp_ComboName[0];
                    temp_FrameInfo.MaxRatio = temp_Ratio[0];

                    // 将杆件结构体加入到对应的点结构体内部的结构体列表中
                    // 分两步，先加入到临时列表中，再将临时列表赋予对应点结构体内部的列表
                    temp_FrameInfoList.Add(temp_FrameInfo);

                    if (FrameType_W_C == 0)
                    {
                        temp_PointInfo.ChordInfoList = temp_FrameInfoList;
                    }
                    else
                    {
                        temp_PointInfo.WebInfoList = temp_FrameInfoList;
                    }
                    

                    // 将修改后的结构体重新放回列表中
                    PointInfoList_1[index] = temp_PointInfo;
                }
                #endregion


                #region //判断point2是否在点结构体列表中，point2在列表中，就输出测站5 ，对应编号4的内力。
                structExists = PointInfoList_1.Exists(s => s.Name == point2);
                if (structExists)
                {
                    //找出另一个点的坐标放入j点信息中
                    //获取点对象的坐标
                    double X = 0;
                    double Y = 0;
                    double Z = 0;
                    ret = SapModel.PointObj.GetCoordCartesian(point1, ref X, ref Y, ref Z);

                    #region 获取杆件设计内力
                    //获取frame的设计内力对应的组合，然后根据这个组合名字来获取内力
                    int temp_NumberItems = 0;
                    string[] temp_FrameName = new string[1];
                    double[] temp_Ratio = new double[1];
                    int[] temp_RatioType = new int[1];
                    double[] temp_Location = new double[1];
                    string[] temp_ComboName = new string[1];
                    string[] temp_ErrorSummary = new string[1];
                    string[] temp_WarningSummary = new string[1];
                    string[] warningSummary = new string[1];

                    ret = SapModel.DesignSteel.GetSummaryResults(frameName, ref temp_NumberItems, ref temp_FrameName, ref temp_Ratio, ref temp_RatioType, ref temp_Location, ref temp_ComboName, ref temp_ErrorSummary, ref temp_WarningSummary, 0);

                    Console.WriteLine(temp_ComboName);
                    //根据这个组合名字来获取内力
                    double[] temp_ItemTypeElm = new double[1];
                    int temp_NumberResults = 0;
                    string[] temp_Obj = new string[1];
                    double[] temp_ObjSta = new double[1];
                    string[] temp_Elm = new string[1];
                    double[] temp_ElmSta = new double[1];
                    string[] temp_LoadCase = temp_ComboName;
                    string[] temp_StepType = new string[1];
                    double[] temp_StepNum = new double[1];
                    double[] temp_P = new double[1];
                    double[] temp_V2 = new double[1];
                    double[] temp_V3 = new double[1];
                    double[] temp_T = new double[1];
                    double[] temp_M2 = new double[1];
                    double[] temp_M3 = new double[1];

                    //frameforce方法来读取
                    //先设置输出组合
                    SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();
                    ret = SapModel.Results.Setup.SetComboSelectedForOutput(temp_ComboName[0], true);

                    ret = SapModel.Results.FrameForce(frameName, 0, ref temp_NumberResults, ref temp_Obj, ref temp_ObjSta, ref temp_Elm, ref temp_ElmSta, ref temp_LoadCase, ref temp_StepType, ref temp_StepNum, ref temp_P, ref temp_V2, ref temp_V3, ref temp_T, ref temp_M2, ref temp_M3);


                    #endregion


                    //寻找point2对应的点结构体
                    temp_PointInfo = PointInfoList_1.Find(s => s.Name == point2);
                    int index = PointInfoList_1.IndexOf(temp_PointInfo);

                    //将结构体中的列表值赋予给临时列表
                    if (FrameType_W_C == 0)
                    {
                        temp_FrameInfoList = temp_PointInfo.ChordInfoList;
                    }
                    else
                    {
                        temp_FrameInfoList = temp_PointInfo.WebInfoList;
                    }

                    //如果temp_PointInfo.WebInfoList;是空的（null），那么需要重新生成下temp_FrameInfoList
                    //因为对null的list进行add，会报错
                    if (temp_FrameInfoList == null)
                    {
                        temp_FrameInfoList = new List<FrameInfo>();
                    }

                    // 修改杆件结构体的成员值
                    temp_FrameInfo.FrameName = frameName;
                    temp_FrameInfo.IName = point2;
                    temp_FrameInfo.JName = point1;
                    temp_FrameInfo.JX = X;
                    temp_FrameInfo.JY = Y;
                    temp_FrameInfo.JZ = Z;
                    temp_FrameInfo.P_max = temp_P.Max();
                    temp_FrameInfo.P_min = temp_P.Min();
                    temp_FrameInfo.P_Cal = (temp_P[4] + temp_P[5]) / 2;
                    temp_FrameInfo.ComboName = temp_ComboName[0];
                    temp_FrameInfo.MaxRatio = temp_Ratio[0];

                    // 将杆件结构体加入到对应的点结构体内部的结构体列表中
                    // 分两步，先加入到临时列表中，再将临时列表赋予对应点结构体内部的列表
                    temp_FrameInfoList.Add(temp_FrameInfo);

                    if (FrameType_W_C == 0)
                    {
                        temp_PointInfo.ChordInfoList = temp_FrameInfoList;
                    }
                    else
                    {
                        temp_PointInfo.WebInfoList = temp_FrameInfoList;
                    }

                    // 将修改后的结构体重新放回列表中
                    PointInfoList_1[index] = temp_PointInfo;
                }
                #endregion


            }
            #endregion
        }


        public struct PointInfo
        {
            public string Name;
            public double pointX;
            public double pointY;
            public double pointZ;
            public List<FrameInfo> ChordInfoList;
            public List<FrameInfo> WebInfoList;
        }
        public struct FrameInfo
        {
            public string FrameName;
            public string IName;
            public string JName;
            public double IX, IY, IZ;
            public double JX, JY, JZ;
            public double P_max, V2, V3, T, M2, M3;
            public double P_min;
            public double P_Cal;
            public string ComboName;
            public double MaxRatio;
        }
    }
}