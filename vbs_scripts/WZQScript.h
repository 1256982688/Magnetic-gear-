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
	//д�ű�
	void CreateScript();
	//���нű�
	void RunScript();
	//���õ�ǰ�ű���Ӧ����Ŀ
	void SetProject();
	//�������ھ��Ե�ַ
	string address;


	//��ͼʱͳһ��Ҫ�Ľű�
	string DrawNeedScripts[2] = { ", Array(\"NAME:Attributes\", \"Name:= \", \"","\", \"Flags:=\", _\n"
		"\"\", \"Color:=\", \"(143 175 143)\", \"Transparency:=\", 0, \"PartCoordinateSystem:=\", _\n"
		"\"Global\", \"UDMId:=\", \"\", \"MaterialValue:=\", \"\" & Chr(34) & "
		"\"vacuum\" & Chr(34) & \"\", \"SurfaceMaterialValue:=\", _\n"
		"\"\" & Chr(34) & \"\" & Chr(34) & \"\", \"SolveInside:=\", true, \"ShellElement:=\", _\n"
		"false, \"ShellElementThickness:=\", \"0mm\", \"IsMaterialEditable:=\", true, "
		"\"UseMaterialAppearance:=\", _\nfalse, \"IsLightweight:=\", false)\n" };

public:
	
	int xx = 0;
	//�Ƿ���Ҫ��������MaxWell
	bool Need = true;
	//project��design��
	string Name = "Maxwell2DDesign1";
	string Project = "Project1";
	~WZQScript()
	{
		cout << "Script" << Project << "��������" << endl;
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
	//����һ���µ���Ŀ
	void CreateProject(/*string name = "Project1"*/);
	//��һ����Ŀ�ڲ���2D��ģ
	void InsertDesign(/*string name = "Maxwell2DDesign1"*/);
	////ɾ������ͼ��
	//void DeleteAll();
	/*2D��ģ-��ͼ*/

	//������
	Rectangle DrawRectangle(string x = "0", string y = "0", string Width = "0.1", string Height = "0.1", string name = "Rectangle1");
	//��Բ
	Circle DrawCircle(string x = "0", string y = "0", string Radius = "0.1", string name = "Circle1");
	//����Բ
	Ellipse DrawEllipse(string x = "0", string y = "0", string MajRadius = "0.1", string Ratio = "0.5", string name = "Ellipse");
	//�������
	Polygon DrawPolygon(string Xcenter = "0", string Ycenter = "0", string Xstart = "0.1", string Ystart = "0.1", string NumSides = "4", string name = "Polygon1");
	//��Բ��
	Arc DrawArc(string startx = "1", string starty = "0", string centerx = "0", string centery = "0", string radian = "90", string name = "Arc1");
	//����
	Polyline DrawPolyline(vector<string> x, vector<string> y, string name = "Polyline1");
	//���ñ߽�
	void SetBoundary();


	/*2D��ģ-�ƶ�ͼ��*/

	//ƽ��
	void move(string name, string x, string y);
	//��ת
	void rotate(string name, string Angle);
	//ƽ������
	void LongLine(string name, string x, string y, string numClones);
	//��ת����
	void AroundAxis(string name, string AngleStr, string numClones);
	//����
	void Mirror(string name, string Basex, string Basey, string Normalx, string Normaly);
	//��������ϵ
	void CreateRelative(string x, string y, string name = "RelativeCS1");
	//�趨��ǰ����ϵ
	void SelectWCS(string name);
	//ɨ��
	void SweepAround(string name, string angle);


	/*���ͼ�β���*/

	//���ñ���
	void SetProperty(string name, string value);
	//���ľ�����״
	void ChangeRectangle(string name, Rectangle R);
	//����Բ����״
	void ChangeCircle(string name, Circle C);
	//���Ķ������״
	void ChangePolygon(string name, Polygon P);
	//������Բ��״
	void ChangeEllipse(string name, Ellipse E);
	//�ϲ��߶γ�ƽ��
	void Connect(string name1, string name2);
	//���Ļ��Ȳ���
	void ChangeArc(string name, Arc A);
	//���ı���ֵ
	void ChangeProperty(string name, string value);
	//Ϊģ�����ñ���
	void SetPropertyforModel(string Target, string name, string property, string p);
	//���ò���
	void ChangeMaterial(string name, string material);
	//������ɫ
	void ChangeColor(string name, int R, int G, int B);
	//���N/S��������
	void NS();


	/*��������*/
	//�г�ͼ��
	void Subtract(string name1, string name2);

	/*��������*/
	//�������ģʽ
	void Transient();
	//�����˶�����
	void SetBand(string name, float rpm);
	//Ϊ�˶���������ת�ٱ���
	void SetBandSpped(string name, string speed);
	//����mesh
	void SetMesh(string name, string target, float MaxLength);
	//���ý������
	void SetAnalysis(string stoptime, string timestep);
	//������Ϣ
	void SaveInformation(string name, string target = "Torque Plot 1");
	//����
	void StartAnalyze();
	//������������
	void CreateReport(string name, string X, string Y);
	//ģ������
	void SetDesignSetting(string l);

	void test();


};


