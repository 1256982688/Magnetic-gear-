#define _CRT_SECURE_NO_WARNINGS
#include "WZQScript.h"
//#include<windows.h>
void WZQScript::test()
{

}
void WZQScript::CreateScript()
{
	//sleep(500);
	fp = fopen(ScriptName.data(), "wt");
	fprintf(fp, "Dim oAnsoftApp\n""Dim oDesktop\n""Dim oProject\n"
		"Dim oDesign\n""Dim oEditor\n""Dim oModule\n"
		"Set oAnsoftApp = CreateObject(\"Ansoft.ElectronicsDesktop\")\n"
		"Set oDesktop = oAnsoftApp.GetAppDesktop()\n"
		"oDesktop.RestoreWindow\n");
}
void WZQScript::RunScript()
{
	string s;
	if (Need)
		s = "wscript.exe " + ScriptName;
	else
		s = ScriptName;
	fclose(fp);
	system(s.data());
	//remove(p);
	struct stat buffer;
	s = Project + ".aedt.lock";
	if (stat(s.c_str(), &buffer) == 0)
		remove(s.c_str());
	s = Project + ".aedt.auto";
	if (stat(s.c_str(), &buffer) == 0)
		remove(s.c_str());
	//remove(p);
}
void WZQScript::SetProject()
{
	fprintf(fp, "Set oProject = oDesktop.SetActiveProject(\"%s\")\n"
		"Set oDesign = oProject.SetActiveDesign(\"%s\")\n"
		"Set oEditor = oDesign.SetActiveEditor(\"3D Modeler\")\n",
		Project.data(), Name.data());
}



//创建一个新的项目
void WZQScript::CreateProject(/*string name*/)
{
	string s = address;
	//Project = name;
	CreateScript();
	fprintf(fp, "Set oProject = oDesktop.NewProject\n"
		"oProject.Rename \"%s/%s.aedt\", true\n", s.data(), Project.data());
	RunScript();
}
//在一个项目内插入2D建模
void WZQScript::InsertDesign(/*string name*/)
{
	//Name = name;
	CreateScript();
	fprintf(fp, "Set oProject = oDesktop.SetActiveProject(\"%s\")\n"
		"oProject.InsertDesign \"Maxwell 2D\", \"%s\", \"Magnetostatic\", \"\"\n", Project.data(), Name.data());
	RunScript();
}



//画矩形
Rectangle WZQScript::DrawRectangle(string x, string y, string Width, string Height, string name)
{

	Rectangle R;
	R.name = name;
	R.Positionx = x, R.Positiony = y, R.Xsize = Width, R.Ysize = Height;
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.CreateRectangle Array(\"NAME:RectangleParameters\", \"IsCovered:=\", true, \"XStart:=\", _\n"
		"\"%s\", \"YStart:=\", \"%s\", \"ZStart:=\", \"0\", \"Width:=\", \"%s\", \"Height:=\", _\n"
		"\"%s\", \"WhichAxis:=\", \"Z\")%s%s%s",
		x.data(), y.data(), Width.data(), Height.data(),
		DrawNeedScripts[0].data(), name.data(), DrawNeedScripts[1].data());
	RunScript();
	return R;

}
//更改矩形参数
void WZQScript::ChangeRectangle(string name, Rectangle R)
{

	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.ChangeProperty Array(\"NAME:AllTabs\", Array(\"NAME:Geometry3DCmdTab\", Array(\"NAME:PropServers\",  _\n\"%s:CreateRectangle:1\"), Array(\"NAME:ChangedProps\", Array(\"NAME:Position\", \"X:=\", _\n\"%s\", \"Y:=\", \"%s\", \"Z:=\", \"0mm\"), Array(\"NAME:XSize\", \"Value:=\", \"%s\"), Array(\"NAME:YSize\", \"Value:=\", _\n\"%s\"))))",
		R.name.data(), R.Positionx.data(), R.Positiony.data(),
		R.Xsize.data(), R.Ysize.data());
	RunScript();

}


//画圆
Circle WZQScript::DrawCircle(string x, string y, string Radius, string name)
{
	Circle C;
	C.name = name;
	C.Positionx = x, C.Positiony = y, C.Radius = Radius;
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.CreateCircle Array(\"NAME:CircleParameters\", \"IsCovered:= \", true, \"XCenter:= \", _\n"
		"\"%s\", \"YCenter:=\", \"%s\", \"ZCenter:=\", \"0mm\", \"Radius:=\", \"%s\", \"WhichAxis:=\", _\n"
		"\"Z\", \"NumSegments:=\", \"0\")%s%s%s",
		x.data(), y.data(), Radius.data(), DrawNeedScripts[0].data(),
		name.data(), DrawNeedScripts[1].data());
	RunScript();
	return C;
}

//更改圆形形状
void WZQScript::ChangeCircle(string name, Circle C)
{
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.ChangeProperty Array(\"NAME:AllTabs\", Array(\"NAME:Geometry3DCmdTab\", Array(\"NAME:PropServers\",  _\n\"%s:CreateCircle:1\"), Array(\"NAME:ChangedProps\", Array(\"NAME:Center Position\", \"\X:=\", _\n\"%s\", \"Y:=\", \"%s\", \"Z:=\", \"0mm\"), Array(\"NAME:Radius\", \"Value:=\", \"%s\"))))",
		name.data(), C.Positionx.data(), C.Positiony.data(),
		C.Radius.data());
	RunScript();
}


//画椭圆
Ellipse WZQScript::DrawEllipse(string x, string y, string MajRadius, string Ratio, string name)
{
	Ellipse E;
	E.name = name;
	E.x = x, E.y = y, E.Radius = MajRadius, E.Ratio = Ratio;
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.CreateEllipse Array(\"NAME:EllipseParameters\", \"IsCovered:=\", true, \"XCenter:=\", _\n\"%s\", \"YCenter:=\", \"%s\", \"ZCenter:=\", \"0mm\", \"MajRadius:=\", \"%s\", \"Ratio:=\", _\n\"%s\", \"WhichAxis:=\", \"Z\", \"NumSegments:=\", \"0\")%s%s%s", x.data(),
		y.data(), MajRadius.data(), Ratio.data(), DrawNeedScripts[0].data(), name.data(), DrawNeedScripts[1].data());
	RunScript();
	return E;
}
//更改椭圆形状
void WZQScript::ChangeEllipse(string name, Ellipse E)
{
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.ChangeProperty Array(\"NAME:AllTabs\", Array(\"NAME:Geometry3DCmdTab\", Array(\"NAME:PropServers\",  _\n\"%s:CreateEllipse:1\"), Array(\"NAME:ChangedProps\", Array(\"NAME:Center Position\", \"X:=\", _\n\"%s\", \"Y:=\", \"%s\", \"Z:=\", \"0mm\"), Array(\"NAME:Major Radius\", \"Value:=\", \"%s\"), Array(\"NAME:Ratio\", \"Value:=\", _\n\"%s\"))))",
		name.data(), E.x.data(), E.y.data(), E.Radius.data(), E.Ratio.data());
	RunScript();
}

//画多边形
Polygon WZQScript::DrawPolygon(string Xcenter, string Ycenter, string Xstart, string Ystart, string NumSides, string name)
{
	Polygon P;
	P.name = name;
	P.Centerx = Xcenter, P.Centery = Ycenter, P.x = Xstart, P.x = Ystart, P.NumberofSegments = NumSides;
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.CreateRegularPolygon Array(\"NAME:RegularPolygonParameters\", \"IsCovered:=\", _\ntrue, \"XCenter:=\", \"%s\", \"YCenter:=\", \"%s\", \"ZCenter:=\", \"0mm\", \"XStart:=\", _\n"
		"\"%s\", \"YStart:=\", \"%s\", \"ZStart:=\", \"0mm\", \"NumSides:=\", \"%s\", \"WhichAxis:=\", _\n"
		"\"Z\")%s%s%s", Xcenter.data(), Ycenter.data(), Xstart.data(),
		Ystart.data(), NumSides.data(), DrawNeedScripts[0].data(), name.data(), DrawNeedScripts[1].data());
	RunScript();
	return P;
}
//更改多边形形状
void WZQScript::ChangePolygon(string name, Polygon P)
{
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.ChangeProperty Array(\"NAME:AllTabs\", Array(\"NAME:Geometry3DCmdTab\", Array(\"NAME:PropServers\",  _\n\"%s:CreateRegularPolygon:1\"), Array(\"NAME:ChangedProps\", Array(\"NAME:Center Position\", \"X:=\", _\n\"%s\", \"Y:=\", \"%s\", \"Z:=\", \"0mm\"), Array(\"NAME:Start Position\", \"X:=\", \"%s\", \"Y:=\", _\n\"%s\", \"Z:=\", \"0mm\"), Array(\"NAME:Number of Segments\", \"Value:=\", \"%s\"))))",
		name.data(), P.Centerx.data(), P.Centery.data(),
		P.x.data(), P.y.data(), P.NumberofSegments.data());
	RunScript();
}

//画折线
Polyline WZQScript::DrawPolyline(vector<string> x, vector<string> y, string name)
{
	Polyline P;
	P.name = name;
	P.x = x;
	P.y = y;
	int l = x.size();
	CreateScript();
	SetProject();
	fprintf(fp, "%s", "oEditor.CreatePolyline Array(\"NAME:PolylineParameters\", \"IsPolylineCovered:= \", true, \"IsPolylineClosed:= \",  false, Array(\"NAME:PolylinePoints\" ");
	for (int i = 0; i < l; i++)
	{
		fprintf(fp, ", Array(\"NAME:PLPoint\", \"X:=\", \"%s\", \"Y:=\", \"%s\", \"Z:=\", \"0mm\")", x[i].data(),
			y[i].data());
	}
	fprintf(fp, "%s", "), Array(\"NAME:PolylineSegments\"");
	for (int i = 0; i < l - 1; i++)
	{
		fprintf(fp, ", Array(\"NAME:PLSegment\", \"SegmentType:=\", \"Line\", \"StartIndex:=\", %d, \"NoOfPoints:=\", 2)", i);
	}
	fprintf(fp, "%s%s%s%s", "), Array(\"NAME:PolylineXSection\", \"XSectionType:=\", \"None\", \"XSectionOrient:=\",\"Auto\", \"XSectionWidth:=\", \"0mm\", \"XSectionTopWidth:=\", \"0mm\", \"XSectionHeight:=\",\"0mm\", \"XSectionNumSegments:=\", \"0\", \"XSectionBendType:=\", \"Corner\"))", DrawNeedScripts[0].data(), name.data(), DrawNeedScripts[1].data());
	RunScript();
	return P;
}
//画圆弧
Arc WZQScript::DrawArc(string startx, string starty, string centerx, string centery, string radian, string name)
{
	Arc A;
	A.name = name;
	A.centerx = centerx, A.centery = centery, A.startx = startx,
		A.starty = starty, A.radian = radian;
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.CreatePolyline Array(\"NAME:PolylineParameters\", \"IsPolylineCovered:=\", true, \"IsPolylineClosed:=\",  _\n"
		"false, Array(\"NAME:PolylinePoints\", Array(\"NAME:PLPoint\", \"X:=\", \"%s\", \"Y:=\", \"%s\", \"Z:=\", _\n"
		"\"0mm\"), Array(\"NAME:PLPoint\", \"X:=\", \"0.4mm\", \"Y:=\", _\n"
		"\"0.4mm\", \"Z:=\", \"0mm\"), Array(\"NAME:PLPoint\", \"X:=\", _\n""\"4.2e-17mm\", \"Y:=\", \"0.7mm\", \"Z:=\", \"0mm\")), Array(\"NAME:PolylineSegments\", Array(\"NAME:PLSegment\", \"SegmentType:=\", _\n""\"AngularArc\", \"StartIndex:=\", 0, \"NoOfPoints:=\", 3, \"NoOfSegments:=\", \"0\", \"ArcAngle:=\", _\n""\"%s\", \"ArcCenterX:=\", \"%s\", \"ArcCenterY:=\", \"%s\", \"ArcCenterZ:=\", \"0mm\", \"ArcPlane:=\", _\n""\"XY\")), Array(\"NAME:PolylineXSection\", \"XSectionType:=\", \"None\", \"XSectionOrient:=\", _\n""\"Auto\", \"XSectionWidth:=\", \"0mm\", \"XSectionTopWidth:=\", \"0mm\", \"XSectionHeight:=\", _\n""\"0mm\", \"XSectionNumSegments:=\", \"0\", \"XSectionBendType:=\", \"Corner\"))%s%s%s",
		startx.data(), starty.data(), radian.data(), centerx.data(),
		centery.data(), DrawNeedScripts[0].data(), name.data(),
		DrawNeedScripts[1].data());
	RunScript();
	return A;
}

//更改弧度参数
void WZQScript::ChangeArc(string name, Arc A)
{
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.ChangeProperty Array(\"NAME:AllTabs\", Array(\"NAME:Geometry3DPolylineTab\", Array(\"NAME:PropServers\",  _\n"
		"\"%s:CreatePolyline:1:Segment0\"), Array(\"NAME:ChangedProps\", Array(\"NAME:Angle\", \"Value:=\", _\n"
		"\"%s\"), Array(\"NAME:Start Point\", \"X:=\", \"%s\", \"Y:=\", \"%s\", \"Z:=\", \"0mm\"))))",
		name.data(), A.radian.data(), A.startx.data(), A.starty.data());
	RunScript();
}
//设置边界
void WZQScript::SetBoundary()
{
	CreateScript();
	SetProject();
	fprintf(fp, "Set oModule = oDesign.GetModule(\"BoundarySetup\")\n"
		"oModule.AssignBalloon Array(\"NAME:Balloon1\", \"Edges:=\", Array(7))");
	RunScript();
}


//平移
void WZQScript::move(string name, string x, string y)
{
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.Move Array(\"NAME:Selections\", \"Selections:=\", \"%s\", \"NewPartsModelFlag:=\", _\n"
		"\"Model\"), Array(\"NAME:TranslateParameters\", \"TranslateVectorX:=\", \"%s\", \"TranslateVectorY:=\", _\n""\"%s\", \"TranslateVectorZ:=\", \"0mm\")\n",
		name.data(), x.data(), y.data());
	RunScript();
}

//旋转
void WZQScript::rotate(string name, string Angle)
{
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.Rotate Array(\"NAME:Selections\", \"Selections:=\", \"%s\", \"NewPartsModelFlag:=\", _\n\"Model\"), Array(\"NAME:RotateParameters\", \"RotateAxis:=\", \"Z\", \"RotateAngle:=\", _\n\"%sdeg\")\n", name.data(), Angle.data());
	RunScript();
}

//平移阵列
void WZQScript::LongLine(string name, string x, string y, string numClones)
{
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.DuplicateAlongLine Array(\"NAME:Selections\", \"Selections:=\", \"%s\", \"NewPartsModelFlag:=\",  _\n"
		"\"Model\"), Array(\"NAME:DuplicateToAlongLineParameters\", \"CreateNewObjects:=\", true, \"XComponent:=\", _\n"
		"\"%s\", \"YComponent:=\", \"%s\", \"ZComponent:=\", \"0mm\", \"NumClones:=\", \"%s\"), Array(\"NAME:Options\", \"DuplicateAssignments:=\", _\ntrue), Array(\"CreateGroupsForNewObjects:=\", false)",
		name.data(), x.data(), y.data(), numClones.data());
	RunScript();
}

//旋转阵列
void WZQScript::AroundAxis(string name, string AngleStr, string numClones)
{
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.DuplicateAroundAxis Array(\"NAME:Selections\", \"Selections:=\", \"%s\", \"NewPartsModelFlag:= \",  _\n"
		"\"Model\"), Array(\"NAME:DuplicateAroundAxisParameters\", \"CreateNewObjects:=\", false, \"WhichAxis:=\", _\n"
		"\"Z\", \"AngleStr:=\", \"%s\", \"NumClones:=\", \"%s\"), Array(\"NAME:Options\", \"DuplicateAssignments:=\", _\n"
		"true), Array(\"CreateGroupsForNewObjects:=\", false)",
		name.data(), AngleStr.data(), numClones.data());
	RunScript();
}
//镜像
void WZQScript::Mirror(string name, string Basex, string Basey, string Normalx, string Normaly)
{
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.Mirror Array(\"NAME:Selections\", \"Selections:= \", \"%s\", \"NewPartsModelFlag:= \",  _\n"
		"\"Model\"), Array(\"NAME:MirrorParameters\", \"MirrorBaseX:=\", \"%s\", \"MirrorBaseY:=\", _\n"
		"\"%s\", \"MirrorBaseZ:=\", \"0mm\", \"MirrorNormalX:=\", \"%s\", \"MirrorNormalY:=\", _\n"
		"\"%s\", \"MirrorNormalZ:=\", \"0mm\")",
		name.data(), Basex.data(), Basey.data(),
		Normalx.data(), Normaly.data());
	RunScript();
}

//扫描
void WZQScript::SweepAround(string name, string angle)
{
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.SweepAroundAxis Array(\"NAME:Selections\", \"Selections:=\", \"%s\", \"NewPartsModelFlag:=\",  _\n"
		"\"Model\"), Array(\"NAME:AxisSweepParameters\", \"DraftAngle:=\", \"0deg\", \"DraftType:=\", _\n"
		"\"Round\", \"CheckFaceFaceIntersection:=\", false, \"SweepAxis:=\", \"Z\", \"SweepAngle:=\", _\n\"%s\", \"NumOfSegments:=\", \"0\")", name.data(), angle.data());
	RunScript();
}

//创造坐标系
void WZQScript::CreateRelative(string x, string y, string name)
{
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.CreateRelativeCS Array(\"NAME:RelativeCSParameters\", \"Mode:= \", \"Axis / Position\", \"OriginX:= \",  _\n"
		"\"%s\", \"OriginY:=\", \"%s\", \"OriginZ:=\", \"0mm\", \"XAxisXvec:=\", \"1mm\", \"XAxisYvec:=\", _\n"
		"\"0mm\", \"XAxisZvec:=\", \"0mm\", \"YAxisXvec:=\", \"0mm\", \"YAxisYvec:=\", \"1mm\", \"YAxisZvec:=\", _\n"
		"\"0mm\"), Array(\"NAME:Attributes\", \"Name:=\", \"%s\")",
		x.data(), y.data(), name.data());
	RunScript();
}
//设定当前坐标系
void WZQScript::SelectWCS(string name)
{
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.SetWCS Array(\"NAME:SetWCS Parameter\", \"Working Coordinate System:=\",  _\n"
		"\"%s\", \"RegionDepCSOk:=\", false)", name.data());
	RunScript();
}

//设置变量
void WZQScript::SetProperty(string name, string value)
{
	CreateScript();
	SetProject();
	fprintf(fp, "oDesign.ChangeProperty Array(\"NAME:AllTabs\", Array(\"NAME:LocalVariableTab\", Array(\"NAME:PropServers\",  _\n"
		"\"LocalVariables\"), Array(\"NAME:NewProps\", Array(\"NAME:%s\", \"PropType:=\", \"VariableProp\", \"UserDef:=\", _\n"
		"true, \"Value:=\", \"%s\"))))", name.data(), value.data());
	RunScript();

}
//更改变量值
void WZQScript::ChangeProperty(string name, string value)
{
	CreateScript();
	SetProject();
	fprintf(fp, "oDesign.ChangeProperty Array(\"NAME:AllTabs\", Array(\"NAME:LocalVariableTab\", Array(\"NAME:PropServers\",  _\n\"LocalVariables\"), Array(\"NAME:ChangedProps\", Array(\"NAME:%s\", \"Value:=\", \"%s\"))))", name.data(), value.data());
	RunScript();
}
//为模型设置变量
void WZQScript::SetPropertyforModel(string Target, string name, string property, string p)
{
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.ChangeProperty Array(\"NAME:AllTabs\", Array(\"NAME:Geometry3D%sTab\", Array(\"NAME:PropServers\",\"%s\"), Array(\"NAME:ChangedProps\", Array(\"NAME:%s\", %s))))",
		p.data(), Target.data(), name.data(), property.data());
	RunScript();
}


//合并线段成平面
void WZQScript::Connect(string name1, string name2)
{
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.Connect Array(\"NAME:Selections\", \"Selections:=\", \"%s, %s\")", name1.data(), name2.data());
	RunScript();
}


//切除图形
void WZQScript::Subtract(string name1, string name2)
{
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.Subtract Array(\"NAME:Selections\", \"Blank Parts:=\", \"%s\", \"Tool Parts:=\",  _\n\"%s\"), Array(\"NAME:SubtractParameters\", \"KeepOriginals:=\", false)", name1.data(), name2.data());
	RunScript();
}

//设置材料
void WZQScript::ChangeMaterial(string name, string material)
{
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.ChangeProperty Array(\"NAME:AllTabs\", Array(\"NAME:Geometry3DAttributeTab\", Array(\"NAME:PropServers\",\"%s\"), Array(\"NAME:ChangedProps\", Array(\"NAME:Material\", \"Value:=\", \"\" & Chr(34) & \"%s\" & Chr(34) & \"\"))))", name.data(), material.data());
	RunScript();
}

//添加N/S极永磁铁
void WZQScript::NS()
{
	CreateScript();
	SetProject();
	fprintf(fp, "Set oDefinitionManager = oProject.GetDefinitionManager()\n"
		"oDefinitionManager.AddMaterial Array(\"NAME:NdFe35_n\", \"CoordinateSystemType:=\", _\n\"Cylindrical\", \"BulkOrSurfaceType:=\", 1, Array(\"NAME:PhysicsTypes\", \"set:=\", Array(_\n\"Electromagnetic\", \"Thermal\", \"Structural\")), Array(\"NAME:AttachedData\", Array(\"NAME:MatAppearanceData\", \"property_data:=\", _\n\"appearance_data\", \"Red:=\", 204, \"Green:=\", 204, \"Blue:=\", 204)), \"permittivity:=\", _\n\"1\", \"permeability:=\", \"1.0997785406\", \"conductivity:=\", \"625000\", \"dielectric_loss_tangent:=\", _\n\"0\", \"magnetic_loss_tangent:=\", \"0\", Array(\"NAME:magnetic_coercivity\", \"property_type:=\", _\n\"VectorProperty\", \"Magnitude:=\", \"-890000A_per_meter\", \"DirComp1:=\", \"1\", \"DirComp2:=\", _\n\"0\", \"DirComp3:=\", \"0\"), \"thermal_conductivity:=\", \"0\", \"saturation_mag:=\", _\n\"0gauss\", \"lande_g_factor:=\", \"2\", \"delta_H:=\", \"0Oe\", \"mass_density:=\", \"7400\", \"youngs_modulus:=\", _\n\"147000000000\", Array(\"NAME:thermal_expansion_coefficient\", \"property_type:=\", _\n\"AnisoProperty\", \"unit:=\", \"\", \"component1:=\", \"3e-06\", \"component2:=\", \"-5e-06\", \"component3:=\", _\n\"-5e-06\"))\n"
		"oDefinitionManager.AddMaterial Array(\"NAME:NdFe35_s\", \"CoordinateSystemType:=\", _\n\"Cylindrical\", \"BulkOrSurfaceType:=\", 1, Array(\"NAME:PhysicsTypes\", \"set:=\", Array(_\n\"Electromagnetic\", \"Thermal\", \"Structural\")), Array(\"NAME:AttachedData\", Array(\"NAME:MatAppearanceData\", \"property_data:=\", _\n\"appearance_data\", \"Red:=\", 204, \"Green:=\", 204, \"Blue:=\", 204)), \"permittivity:=\", _\n\"1\", \"permeability:=\", \"1.0997785406\", \"conductivity:=\", \"625000\", \"dielectric_loss_tangent:=\", _\n\"0\", \"magnetic_loss_tangent:=\", \"0\", Array(\"NAME:magnetic_coercivity\", \"property_type:=\", _\n\"VectorProperty\", \"Magnitude:=\", \"-890000A_per_meter\", \"DirComp1:=\", \"-1\", \"DirComp2:=\", _\n\"0\", \"DirComp3:=\", \"0\"), \"thermal_conductivity:=\", \"0\", \"saturation_mag:=\", _\n\"0gauss\", \"lande_g_factor:=\", \"2\", \"delta_H:=\", \"0Oe\", \"mass_density:=\", \"7400\", \"youngs_modulus:=\", _\n\"147000000000\", Array(\"NAME:thermal_expansion_coefficient\", \"property_type:=\", _\n\"AnisoProperty\", \"unit:=\", \"\", \"component1:=\", \"3e-06\", \"component2:=\", \"-5e-06\", \"component3:=\", _\n\"-5e-06\"))");
	RunScript();
}

//设置颜色
void WZQScript::ChangeColor(string name, int R, int G, int B)
{
	CreateScript();
	SetProject();
	fprintf(fp, "oEditor.ChangeProperty Array(\"NAME:AllTabs\", Array(\"NAME:Geometry3DAttributeTab\", Array(\"NAME:PropServers\",  _\n"
		"\"%s\"), Array(\"NAME:ChangedProps\", Array(\"NAME:Color\", \"R:=\", %d, \"G:=\", %d, \"B:=\", %d))))", name.data(), R, G, B);
	RunScript();
}

//设置求解模式
void WZQScript::Transient()
{
	CreateScript();
	SetProject();
	fprintf(fp, "oDesign.SetSolutionType \"Transient\", \"XY\"");
	RunScript();
}

//设置运动区域
void WZQScript::SetBand(string name, float rpm)
{
	CreateScript();
	SetProject();
	fprintf(fp, "Set oModule = oDesign.GetModule(\"ModelSetup\")\n"
		"oModule.AssignBand Array(\"NAME:Data\", \"Move Type:=\", \"Rotate\", \"Coordinate System:=\","
		"  _\n\"Global\", \"Axis:=\", \"Z\", \"Is Positive:=\", true, "
		"\"InitPos:=\", \"0deg\", \"HasRotateLimit:=\", _\nfalse, \"NonCylindrical:=\", "
		"false, \"Consider Mechanical Transient:=\", false, \"Angular Velocity:=\", _\n\"%frpm\","
		" \"Objects:=\", Array(\"%s\"))"
		, rpm, name.data());
	RunScript();
}

//为运动区域设置转速变量
void WZQScript::SetBandSpped(string name, string speed)
{
	CreateScript();
	SetProject();
	fprintf(fp, "Set oModule = oDesign.GetModule(\"ModelSetup\")\n""oModule.EditMotionSetup \"%s\", Array(\"NAME:Data\", \"Move Type:=\", \"Rotate\", \"Coordinate System:=\",  _\n"
		"\"Global\", \"Axis:=\", \"Z\", \"Is Positive:=\", true, \"InitPos:=\", \"0deg\", \"HasRotateLimit:=\", _\n"
		"false, \"NonCylindrical:=\", false, \"Consider Mechanical Transient:=\", false, \"Angular Velocity:=\", _\n"
		"\"%srpm\")\n", name.data(), speed.data());
	RunScript();
}

//设置mesh
void WZQScript::SetMesh(string name, string target, float MaxLength)
{
	CreateScript();
	SetProject();
	fprintf(fp, "Set oModule = oDesign.GetModule(\"MeshSetup\")\n"
		"oModule.AssignLengthOp Array(\"NAME:%s\", \"RefineInside:=\", false, \"Enabled:=\",  _\ntrue, \"Objects:=\", Array(\"%s\"), \"RestrictElem:=\", false, \"NumMaxElem:=\", _\n\"1000\", \"RestrictLength:=\", true, \"MaxLength:=\", \"%fmm\")",
		name.data(), target.data(), MaxLength);
	RunScript();
}

//设置解决方案
void WZQScript::SetAnalysis(string stoptime,string timestep)
{
	CreateScript();
	SetProject();
	fprintf(fp,"Set oModule = oDesign.GetModule(\"AnalysisSetup\")\n"
	"oModule.InsertSetup \"Transient\", Array(\"NAME:Setup1\", \"Enabled:=\", true, Array(\"NAME:MeshLink\", \"ImportMesh:=\",  _\n"
	"  false), \"NonlinearSolverResidual:=\", \"0.0001\", \"TimeIntegrationMethod:=\",  _\n"
	"  \"BackwardEuler\", \"SmoothBHCurve:=\", false, \"StopTime:=\", \"%s\", \"TimeStep:=\",  _\n"
	" \"%s\", \"OutputError:=\", false, \"UseControlProgram:=\", false, \"ControlProgramName:=\",  _\n"
	" \"\", \"ControlProgramArg:=\", \" \", \"CallCtrlProgAfterLastStep:=\", false, \"FastReachSteadyState:=\",  _\n"
	"  false, \"AutoDetectSteadyState:=\", false, \"IsGeneralTransient:=\", true, \"IsHalfPeriodicTransient:=\",  _\n"
	"  false, \"SaveFieldsType:=\", \"None\", \"UseNonLinearIterNum:=\", false, \"CacheSaveKind:=\",  _\n"
	"  \"Count\", \"NumberSolveSteps:=\", 1, \"RangeStart:=\", \"0s\", \"RangeEnd:=\", \"0.1s\", \"UseAdaptiveTimeStep:=\",  _\n"
	"  false, \"InitialTimeStep:=\", \"0.002s\", \"MinTimeStep:=\", \"0.001s\", \"MaxTimeStep:=\",  _\n"
	"  \"0.003s\", \"TimeStepErrTolerance:=\", 0.0001)",stoptime.data(),timestep.data());
	RunScript();
}

//保存信息
void WZQScript::SaveInformation(string name, string target)
{
	//Torque Plot 1
	for (int i = 0; i < address.size(); i++)
	{
		if (address[i] == '\\')
			address[i] = '/';
	}
	name = address + "/" + name;
	CreateScript();
	SetProject();
	fprintf(fp, "Set oModule = oDesign.GetModule(\"ReportSetup\")\n"
		"oModule.ExportToFile \"%s\", _\n"
		"\"%s\", false", target.data(), name.data());
	RunScript();
}

//分析
void WZQScript::StartAnalyze()
{
	CreateScript();
	SetProject();
	fprintf(fp, "Set oProject = oDesktop.SetActiveProject(\"%s\")\n"
		"oProject.Save\n"
		"Set oDesign = oProject.SetActiveDesign(\"%s\")\n"
		"oDesign.AnalyzeAll\n", Project.data(), Name.data());
	RunScript();
}

//创建分析报告
//X=Time    Y=\"Moving.Torque\",\"Moving.Torque\"
void WZQScript::CreateReport(string name, string X, string Y)
{
	CreateScript();
	SetProject();
	fprintf(fp, "Set oModule = oDesign.GetModule(\"ReportSetup\")\n"
		"oModule.CreateReport \"%s\", \"Transient\", \"Rectangular Plot\", _\n"
		"\"Setup1 : Transient\", Array(\"Domain:=\", \"Sweep\"), Array(\"Time:=\","
		" Array(\"All\")), Array(\"X Component:=\", \"%s\", \"Y Component:=\", Array(_\n"
		"%s))", name.data(), X.data(), Y.data());
	RunScript();
}

//模型设置
void WZQScript::SetDesignSetting(string l)
{
	CreateScript();
	SetProject();
	fprintf(fp, "oDesign.SetDesignSettings Array(\"NAME:Design Settings Data\", \"Perform Minimal validation:=\",  _\n"
		"false, \"EnabledObjects:=\", Array(), \"PreserveTranSolnAfterDatasetEdit:=\", _\n"
		"false, \"ComputeTransientInductance:=\", false, \"ComputeIncrementalMatrix:=\", _\n"
		"false, \"PerfectConductorThreshold:=\", 1E+30, \"InsulatorThreshold:=\", 1, \"ModelDepth:=\", _\n"
		"\"%s\", \"UseSkewModel:=\", false, \"EnableTranTranLinkWithSimplorer:=\", false, \"BackgroundMaterialName:=\", _\n"
		"\"vacuum\", \"SolveFraction:=\", false, \"Multiplier:=\", \"1\"), Array(\"NAME:Model Validation Settings\", \"EntityCheckLevel:=\", _\n"
		"\"Strict\", \"IgnoreUnclassifiedObjects:=\", false, \"SkipIntersectionChecks:=\", _\nfalse)", l.data());
	RunScript();
}