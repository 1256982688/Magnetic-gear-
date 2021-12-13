#pragma once
#include<iostream>
#include<map>
#include<vector>
#include<string>
#include<fstream>
#include <direct.h>
#include <sys/stat.h>
//#include <unistd.h>
//#include<Windows.h>

using namespace std;
struct Rectangle
{
	string name;
	string Positionx;
	string Positiony;
	string Xsize;
	string Ysize;
};
struct Circle
{
	string name;
	string Positionx;
	string Positiony;
	string Radius;
};
struct Polygon
{
	string name;
	string Centerx;
	string Centery;
	string x;
	string y;
	string NumberofSegments;
};
struct Ellipse
{
	string name;
	string x;
	string y;
	string Radius;
	string Ratio;
};
struct Arc
{
	string name;
	string centerx;
	string centery;
	string startx;
	string starty;
	string radian;
};
struct Polyline
{
	string name;
	vector<string> x;
	vector<string> y;
};

class WZQScript
{
private:
	string ScriptName = "Script.vbs";
	const char* p = ScriptName.data();
	FILE* fp = new FILE;
	//写脚本
	void CreateScript();
	//运行脚本
	void RunScript();
	//设置当前脚本对应的项目
	void SetProject();
	//程序所在绝对地址
	string address;


	//画图时统一需要的脚本
	string DrawNeedScripts[2] = { ", Array(\"NAME:Attributes\", \"Name:= \", \"","\", \"Flags:=\", _\n"
		"\"\", \"Color:=\", \"(143 175 143)\", \"Transparency:=\", 0, \"PartCoordinateSystem:=\", _\n"
		"\"Global\", \"UDMId:=\", \"\", \"MaterialValue:=\", \"\" & Chr(34) & "
		"\"vacuum\" & Chr(34) & \"\", \"SurfaceMaterialValue:=\", _\n"
		"\"\" & Chr(34) & \"\" & Chr(34) & \"\", \"SolveInside:=\", true, \"ShellElement:=\", _\n"
		"false, \"ShellElementThickness:=\", \"0mm\", \"IsMaterialEditable:=\", true, "
		"\"UseMaterialAppearance:=\", _\nfalse, \"IsLightweight:=\", false)\n" };

public:
	
	int xx = 0;
	//是否需要重新启动MaxWell
	bool Need = true;
	//project和design名
	string Name = "Maxwell2DDesign1";
	string Project = "Project1";
	~WZQScript()
	{
		cout << "Script" << Project << "析构函数" << endl;
		delete fp;
	}
	WZQScript()
	{
		char buff[256];
		_getcwd(buff, sizeof(buff));
		string s = "";
		for (int i = 0; i < sizeof(buff); i++)
		{
			if (buff[i] == '\0')
				break;
			s += buff[i];
		}
		address = s;
	}
	WZQScript(string projectname, string designname, string vbsname)
	{
		Project = projectname;
		Name = designname;
		ScriptName = vbsname;
		char buff[256];
		_getcwd(buff, sizeof(buff));
		string s = "";
		for (int i = 0; i < sizeof(buff); i++)
		{
			if (buff[i] == '\0')
				break;
			s += buff[i];
		}
		address = s;
	}
	//创建一个新的项目
	void CreateProject(/*string name = "Project1"*/);
	//在一个项目内插入2D建模
	void InsertDesign(/*string name = "Maxwell2DDesign1"*/);
	////删除所有图形
	//void DeleteAll();
	/*2D建模-画图*/

	//画矩形
	Rectangle DrawRectangle(string x = "0", string y = "0", string Width = "0.1", string Height = "0.1", string name = "Rectangle1");
	//画圆
	Circle DrawCircle(string x = "0", string y = "0", string Radius = "0.1", string name = "Circle1");
	//画椭圆
	Ellipse DrawEllipse(string x = "0", string y = "0", string MajRadius = "0.1", string Ratio = "0.5", string name = "Ellipse");
	//画多边形
	Polygon DrawPolygon(string Xcenter = "0", string Ycenter = "0", string Xstart = "0.1", string Ystart = "0.1", string NumSides = "4", string name = "Polygon1");
	//画圆弧
	Arc DrawArc(string startx = "1", string starty = "0", string centerx = "0", string centery = "0", string radian = "90", string name = "Arc1");
	//折线
	Polyline DrawPolyline(vector<string> x, vector<string> y, string name = "Polyline1");
	//设置边界
	void SetBoundary();


	/*2D建模-移动图形*/

	//平移
	void move(string name, string x, string y);
	//旋转
	void rotate(string name, string Angle);
	//平移阵列
	void LongLine(string name, string x, string y, string numClones);
	//旋转阵列
	void AroundAxis(string name, string AngleStr, string numClones);
	//镜像
	void Mirror(string name, string Basex, string Basey, string Normalx, string Normaly);
	//创造坐标系
	void CreateRelative(string x, string y, string name = "RelativeCS1");
	//设定当前坐标系
	void SelectWCS(string name);
	//扫描
	void SweepAround(string name, string angle);


	/*变更图形参数*/

	//设置变量
	void SetProperty(string name, string value);
	//更改矩形形状
	void ChangeRectangle(string name, Rectangle R);
	//更改圆形形状
	void ChangeCircle(string name, Circle C);
	//更改多边形形状
	void ChangePolygon(string name, Polygon P);
	//更改椭圆形状
	void ChangeEllipse(string name, Ellipse E);
	//合并线段成平面
	void Connect(string name1, string name2);
	//更改弧度参数
	void ChangeArc(string name, Arc A);
	//更改变量值
	void ChangeProperty(string name, string value);
	//为模型设置变量
	void SetPropertyforModel(string Target, string name, string property, string p);
	//设置材料
	void ChangeMaterial(string name, string material);
	//设置颜色
	void ChangeColor(string name, int R, int G, int B);
	//添加N/S极永磁铁
	void NS();


	/*布尔运算*/
	//切除图形
	void Subtract(string name1, string name2);

	/*运算需求*/
	//设置求解模式
	void Transient();
	//设置运动区域
	void SetBand(string name, float rpm);
	//为运动区域设置转速变量
	void SetBandSpped(string name, string speed);
	//设置mesh
	void SetMesh(string name, string target, float MaxLength);
	//设置解决方案
	void SetAnalysis(string stoptime, string timestep);
	//保存信息
	void SaveInformation(string name, string target = "Torque Plot 1");
	//分析
	void StartAnalyze();
	//创建分析报告
	void CreateReport(string name, string X, string Y);
	//模型设置
	void SetDesignSetting(string l);

	void test();


};


