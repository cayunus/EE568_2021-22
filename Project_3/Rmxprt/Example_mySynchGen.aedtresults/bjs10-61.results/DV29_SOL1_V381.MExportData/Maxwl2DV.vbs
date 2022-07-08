' ----------------------------------------------------------------------
' Script Created by RMxprt to generate Maxwell 2D Version 2016.0 design 
' Can specify one arg to setup external circuit                         
' ----------------------------------------------------------------------

On Error Resume Next
Set oAnsoftApp = CreateObject("Ansoft.ElectronicsDesktop")
On Error Goto 0
Set oDesktop = oAnsoftApp.GetAppDesktop()
Set oArgs = AnsoftScript.arguments
oDesktop.RestoreWindow
Set oProject = oDesktop.GetActiveProject()
Set oDesign = oProject.GetActiveDesign()
designName = oDesign.GetName
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.SetModelUnits Array("NAME:Units Parameter", "Units:=", "mm", _
  "Rescale:=", false)
oDesign.SetSolutionType "Transient", "XY"
Set oModule = oDesign.GetModule("BoundarySetup")
if (oArgs.Count = 1) then 
oModule.EditExternalCircuit oArgs(0), Array(), Array(), Array(), Array()
end if
oEditor.SetModelValidationSettings Array("NAME:Validation Options", _
  "EntityCheckLevel:=", "Strict", "IgnoreUnclassifiedObjects:=", true)
On Error Resume Next
Set oModule = oDesign.GetModule("MeshSetup")
oModule.InitialMeshSettings Array("NAME:MeshSettings", "MeshMethod:=", _
  "AnsoftTAU")
On Error Goto 0
On Error Resume Next
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:fractions", "PropType:=", "VariableProp", "UserDef:=", true, _
  "Value:=", "4"))))
On Error Goto 0
Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.ModifyLibraries designName, Array("NAME:PersonalLib"), _
  Array("NAME:UserLib"), Array("NAME:SystemLib", "Materials:=", Array(_
  "Materials", "RMxprt"))
if (oDefinitionManager.DoesMaterialExist("M22_24G_2DSF0.826")) then
else
oDefinitionManager.AddMaterial Array("NAME:M22_24G_2DSF0.826", _
  "CoordinateSystemType:=", "Cartesian", Array("NAME:AttachedData"), Array(_
  "NAME:ModifierData"), Array("NAME:permeability", "property_type:=", _
  "nonlinear", "HUnit:=", "A_per_meter", "BUnit:=", "tesla", Array(_
  "NAME:BHCoordinates", Array("NAME:Coordinate", "X:=", 0, "Y:=", 0), Array(_
  "NAME:Coordinate", "X:=", 160, "Y:=", 0.661197), Array("NAME:Coordinate", _
  "X:=", 240, "Y:=", 0.826505), Array("NAME:Coordinate", "X:=", 320, "Y:=", _
  0.909167), Array("NAME:Coordinate", "X:=", 420, "Y:=", 0.991835), Array(_
  "NAME:Coordinate", "X:=", 680, "Y:=", 1.07454), Array("NAME:Coordinate", _
  "X:=", 810, "Y:=", 1.09523), Array("NAME:Coordinate", "X:=", 1100, "Y:=", _
  1.13661), Array("NAME:Coordinate", "X:=", 1400, "Y:=", 1.15734), Array(_
  "NAME:Coordinate", "X:=", 2600, "Y:=", 1.21958), Array("NAME:Coordinate", _
  "X:=", 4200, "Y:=", 1.27365), Array("NAME:Coordinate", "X:=", 5087.91, _
  "Y:=", 1.30277), Array("NAME:Coordinate", "X:=", 6579.41, "Y:=", 1.34442), _
  Array("NAME:Coordinate", "X:=", 8231.91, "Y:=", 1.3861), Array(_
  "NAME:Coordinate", "X:=", 10258.9, "Y:=", 1.42787), Array("NAME:Coordinate", _
  "X:=", 13348.9, "Y:=", 1.46986), Array("NAME:Coordinate", "X:=", 18788.9, _
  "Y:=", 1.51237), Array("NAME:Coordinate", "X:=", 28908.9, "Y:=", 1.5559), _
  Array("NAME:Coordinate", "X:=", 46508.9, "Y:=", 1.60106), Array(_
  "NAME:Coordinate", "X:=", 74308.9, "Y:=", 1.64845), Array("NAME:Coordinate", _
  "X:=", 129350, "Y:=", 1.72244), Array("NAME:Coordinate", "X:=", 248716, _
  "Y:=", 1.87244), Array("NAME:Coordinate", "X:=", 447658, "Y:=", 2.12244), _
  Array("NAME:Coordinate", "X:=", 726178, "Y:=", 2.47244), Array(_
  "NAME:Coordinate", "X:=", 3.51138e+06, "Y:=", 5.97243))), Array(_
  "NAME:core_loss_type", "property_type:=", "ChoiceProperty", "Choice:=", _
  "Electrical Steel"), "core_loss_kh:=", 215.015, "core_loss_kc:=", 1.60929, _
  "core_loss_ke:=", 1.342, "conductivity:=", 2e+06, "mass_density:=", 6322.36) 
end if
if (oDefinitionManager.DoesMaterialExist("M22_24G_2DSF0.999")) then
else
oDefinitionManager.AddMaterial Array("NAME:M22_24G_2DSF0.999", _
  "CoordinateSystemType:=", "Cartesian", Array("NAME:AttachedData"), Array(_
  "NAME:ModifierData"), Array("NAME:permeability", "property_type:=", _
  "nonlinear", "HUnit:=", "A_per_meter", "BUnit:=", "tesla", Array(_
  "NAME:BHCoordinates", Array("NAME:Coordinate", "X:=", 0, "Y:=", 0), Array(_
  "NAME:Coordinate", "X:=", 160, "Y:=", 0.798933), Array("NAME:Coordinate", _
  "X:=", 240, "Y:=", 0.998666), Array("NAME:Coordinate", "X:=", 320, "Y:=", _
  1.09853), Array("NAME:Coordinate", "X:=", 420, "Y:=", 1.1984), Array(_
  "NAME:Coordinate", "X:=", 680, "Y:=", 1.29827), Array("NAME:Coordinate", _
  "X:=", 810, "Y:=", 1.32323), Array("NAME:Coordinate", "X:=", 1100, "Y:=", _
  1.37317), Array("NAME:Coordinate", "X:=", 1400, "Y:=", 1.39813), Array(_
  "NAME:Coordinate", "X:=", 2600, "Y:=", 1.47304), Array("NAME:Coordinate", _
  "X:=", 4200, "Y:=", 1.53795), Array("NAME:Coordinate", "X:=", 5087.91, _
  "Y:=", 1.57291), Array("NAME:Coordinate", "X:=", 6579.41, "Y:=", 1.62284), _
  Array("NAME:Coordinate", "X:=", 8231.91, "Y:=", 1.67278), Array(_
  "NAME:Coordinate", "X:=", 10258.9, "Y:=", 1.72272), Array("NAME:Coordinate", _
  "X:=", 13348.9, "Y:=", 1.77265), Array("NAME:Coordinate", "X:=", 18788.9, _
  "Y:=", 1.8226), Array("NAME:Coordinate", "X:=", 28908.9, "Y:=", 1.87255), _
  Array("NAME:Coordinate", "X:=", 46508.9, "Y:=", 1.92251), Array(_
  "NAME:Coordinate", "X:=", 74308.9, "Y:=", 1.97249), Array("NAME:Coordinate", _
  "X:=", 129350, "Y:=", 2.04748), Array("NAME:Coordinate", "X:=", 248716, _
  "Y:=", 2.19748), Array("NAME:Coordinate", "X:=", 447658, "Y:=", 2.44748), _
  Array("NAME:Coordinate", "X:=", 726178, "Y:=", 2.79748), Array(_
  "NAME:Coordinate", "X:=", 3.51138e+06, "Y:=", 6.29748))), Array(_
  "NAME:core_loss_type", "property_type:=", "ChoiceProperty", "Choice:=", _
  "Electrical Steel"), "core_loss_kh:=", 177.937, "core_loss_kc:=", 1.33178, _
  "core_loss_ke:=", 1.22081, "conductivity:=", 2e+06, "mass_density:=", _
  7639.79) 
end if
if (oDefinitionManager.DoesMaterialExist("copper_100C")) then
else
oDefinitionManager.AddMaterial Array("NAME:copper_100C", _
  "CoordinateSystemType:=", "Cartesian", Array("NAME:AttachedData"), Array(_
  "NAME:ModifierData"), "conductivity:=", "4.25672e+07")
end if
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/Band", "Version:=", "12.1", "NoOfParameters:=", 7, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5414.5mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "1800mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", _
  "0deg"), Array("NAME:Pair", "Name:=", "Fractions", "Value:=", "1"), Array(_
  "NAME:Pair", "Name:=", "HalfAxial", "Value:=", "0"), Array("NAME:Pair", _
  "Name:=", "InfoCore", "Value:=", "0"))), Array("NAME:Attributes", "Name:=", _
  "Band", "Flags:=", "", "Color:=", "(0 255 255)", "Transparency:=", 0, _
  "PartCoordinateSystem:=", "Global", "MaterialName:=", "vacuum", _
  "SolveInside:=", true) 
On Error Resume Next
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/Band", "Version:=", "12.1", "NoOfParameters:=", 7, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5414.5mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "1800mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", _
  "0deg"), Array("NAME:Pair", "Name:=", "Fractions", "Value:=", "1"), Array(_
  "NAME:Pair", "Name:=", "HalfAxial", "Value:=", "0"), Array("NAME:Pair", _
  "Name:=", "InfoCore", "Value:=", "100"))), Array("NAME:Attributes", _
  "Name:=", "Shaft", "Flags:=", "", "Color:=", "(0 255 255)", _
  "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "MaterialName:=", _
  "vacuum", "SolveInside:=", true) 
On Error Resume Next
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/Band", "Version:=", "12.1", "NoOfParameters:=", 7, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5414.5mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "5974mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", _
  "0deg"), Array("NAME:Pair", "Name:=", "Fractions", "Value:=", "4"), Array(_
  "NAME:Pair", "Name:=", "HalfAxial", "Value:=", "0"), Array("NAME:Pair", _
  "Name:=", "InfoCore", "Value:=", "100"))), Array("NAME:Attributes", _
  "Name:=", "OuterRegion", "Flags:=", "", "Color:=", "(0 255 255)", _
  "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "MaterialName:=", _
  "vacuum", "SolveInside:=", true) 
On Error Resume Next
On Error Goto 0
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion:CreateUserDefinedPart:1", "Fractions", "fractions"
oEditor.Copy Array("NAME:Selections", "Selections:=", "OuterRegion")
oEditor.Paste
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion1:CreateUserDefinedPart:1", "InfoCore", "1"
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "OuterRegion1"), _
  Array("NAME:ChangedProps", Array("NAME:Name", "Value:=", "Tool"))))
oEditor.FitAll
Set oModule = oDesign.GetModule("ModelSetup")
oModule.SetSymmetryMultiplier "fractions"
Set oModule = oDesign.GetModule("BoundarySetup")
edgeID = oEditor.GetEdgeByPosition(Array("NAME:Parameters", "BodyName:=", _
  "OuterRegion", "XPosition:=", "2112.1279554042176mm", "YPosition:=", _
  "2112.1279554042176mm", "ZPosition:=", "0mm"))
oModule.AssignVectorPotential Array("NAME:VectorPotential1", "Edges:=", Array(edgeID), _
  "Value:=", "0", "CoordinateSystem:=", "")
edgeID = oEditor.GetEdgeByPosition(Array("NAME:Parameters", "BodyName:=", _
  "OuterRegion", "XPosition:=", "1493.5mm", "YPosition:=", "0mm", _
  "ZPosition:=", "0mm"))
oModule.AssignMaster Array("NAME:Master1", "Edges:=", Array(edgeID), _
  "ReverseV:=", false)
edgeID = oEditor.GetEdgeByPosition(Array("NAME:Parameters", "BodyName:=", _
  "OuterRegion", "XPosition:=", "9.1450499726328604e-14mm", "YPosition:=", _
  "1493.5mm", "ZPosition:=", "0mm"))
oModule.AssignSlave Array("NAME:Slave1", "Edges:=", Array(edgeID), _
  "ReverseU:=", true, "Master:=", "Master1", "SameAsMaster:=", true)
oDesign.SetDesignSettings Array("NAME:Design Settings Data", "ModelDepth:=", _
  "1057.41mm")
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SlotCore", "Version:=", "12.1", "NoOfParameters:=", 19, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5431mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "5974mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", "Value:=", "0deg"), _
  Array("NAME:Pair", "Name:=", "Slots", "Value:=", "300"), Array("NAME:Pair", _
  "Name:=", "SlotType", "Value:=", "6"), Array("NAME:Pair", "Name:=", "Hs0", _
  "Value:=", "1mm"), Array("NAME:Pair", "Name:=", "Hs01", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "4.2mm"), Array("NAME:Pair", _
  "Name:=", "Hs2", "Value:=", "120mm"), Array("NAME:Pair", "Name:=", "Bs0", _
  "Value:=", "22.7mm"), Array("NAME:Pair", "Name:=", "Bs1", "Value:=", _
  "22.7mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", "22.7mm"), Array(_
  "NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", "HalfSlot", _
  "Value:=", "0"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", _
  "15deg"), Array("NAME:Pair", "Name:=", "LenRegion", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", "0"))), Array(_
  "NAME:Attributes", "Name:=", "Stator", "Flags:=", "", "Color:=", _
  "(132 132 193)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "M22_24G_2DSF0.826", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Stator:CreateUserDefinedPart:1", _
  "LenRegion", "1180mm + 2*endRegion"
On Error Goto 0
On Error Resume Next
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:delta", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "-25.908893038327076deg"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:conds", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "1"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:R1", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "0.013834093791735462ohm"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:Le1", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "0.00046156375900369416H"))))
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/LapCoil", "Version:=", "12.0", "NoOfParameters:=", 21, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5431mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "5974mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", "Value:=", "0deg"), _
  Array("NAME:Pair", "Name:=", "Slots", "Value:=", "300"), Array("NAME:Pair", _
  "Name:=", "SlotType", "Value:=", "6"), Array("NAME:Pair", "Name:=", "Hs0", _
  "Value:=", "1mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "4.2mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "120mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "22.7mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "22.7mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", _
  "22.7mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", _
  "Name:=", "Layers", "Value:=", "2"), Array("NAME:Pair", "Name:=", _
  "CoilPitch", "Value:=", "9"), Array("NAME:Pair", "Name:=", "EndExt", _
  "Value:=", "150mm"), Array("NAME:Pair", "Name:=", "SpanExt", "Value:=", _
  "0mm"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", "10deg"), _
  Array("NAME:Pair", "Name:=", "LenRegion", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "InfoCoil", "Value:=", "2"))), Array(_
  "NAME:Attributes", "Name:=", "Coil", "Flags:=", "", "Color:=", _
  "(250 167 14)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "copper", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Coil:CreateUserDefinedPart:1", _
  "LenRegion", "1180mm + 2*endRegion"
On Error Goto 0
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "Coil"), _
  Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, _
  "CreateNewObjects:=", true, "WhichAxis:=", "Z", "AngleStr:=", "1.2deg", _
  "NumClones:=", "300"), Array("NAME:Options", "DuplicateBoundaries:=", _
  false)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "Coil"), Array(_
  "NAME:ChangedProps", Array("NAME:Name", "Value:=", "Coil_0"))))
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_0,Coil_75,Coil_150" & _
  ",Coil_225"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_1,Coil_76,Coil_151" & _
  ",Coil_226"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_2,Coil_77,Coil_152" & _
  ",Coil_227"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_3,Coil_78,Coil_153" & _
  ",Coil_228"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_4,Coil_79,Coil_154" & _
  ",Coil_229"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_5,Coil_80,Coil_155" & _
  ",Coil_230"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_6,Coil_81,Coil_156" & _
  ",Coil_231"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_7,Coil_82,Coil_157" & _
  ",Coil_232"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_8,Coil_83,Coil_158" & _
  ",Coil_233"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_9,Coil_84,Coil_159" & _
  ",Coil_234"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_10,Coil_85" & _
  ",Coil_160,Coil_235"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_11,Coil_86" & _
  ",Coil_161,Coil_236"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_12,Coil_87" & _
  ",Coil_162,Coil_237"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_13,Coil_88" & _
  ",Coil_163,Coil_238"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_14,Coil_89" & _
  ",Coil_164,Coil_239"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_15,Coil_90" & _
  ",Coil_165,Coil_240"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_16,Coil_91" & _
  ",Coil_166,Coil_241"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_17,Coil_92" & _
  ",Coil_167,Coil_242"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_18,Coil_93" & _
  ",Coil_168,Coil_243"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_19,Coil_94" & _
  ",Coil_169,Coil_244"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_20,Coil_95" & _
  ",Coil_170,Coil_245"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_21,Coil_96" & _
  ",Coil_171,Coil_246"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_22,Coil_97" & _
  ",Coil_172,Coil_247"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_23,Coil_98" & _
  ",Coil_173,Coil_248"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_24,Coil_99" & _
  ",Coil_174,Coil_249"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_25,Coil_100" & _
  ",Coil_175,Coil_250"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_26,Coil_101" & _
  ",Coil_176,Coil_251"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_27,Coil_102" & _
  ",Coil_177,Coil_252"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_28,Coil_103" & _
  ",Coil_178,Coil_253"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_29,Coil_104" & _
  ",Coil_179,Coil_254"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_30,Coil_105" & _
  ",Coil_180,Coil_255"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_31,Coil_106" & _
  ",Coil_181,Coil_256"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_32,Coil_107" & _
  ",Coil_182,Coil_257"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_33,Coil_108" & _
  ",Coil_183,Coil_258"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_34,Coil_109" & _
  ",Coil_184,Coil_259"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_35,Coil_110" & _
  ",Coil_185,Coil_260"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_36,Coil_111" & _
  ",Coil_186,Coil_261"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_37,Coil_112" & _
  ",Coil_187,Coil_262"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_38,Coil_113" & _
  ",Coil_188,Coil_263"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_39,Coil_114" & _
  ",Coil_189,Coil_264"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_40,Coil_115" & _
  ",Coil_190,Coil_265"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_41,Coil_116" & _
  ",Coil_191,Coil_266"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_42,Coil_117" & _
  ",Coil_192,Coil_267"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_43,Coil_118" & _
  ",Coil_193,Coil_268"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_44,Coil_119" & _
  ",Coil_194,Coil_269"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_45,Coil_120" & _
  ",Coil_195,Coil_270"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_46,Coil_121" & _
  ",Coil_196,Coil_271"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_47,Coil_122" & _
  ",Coil_197,Coil_272"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_48,Coil_123" & _
  ",Coil_198,Coil_273"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_49,Coil_124" & _
  ",Coil_199,Coil_274"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_50,Coil_125" & _
  ",Coil_200,Coil_275"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_51,Coil_126" & _
  ",Coil_201,Coil_276"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_52,Coil_127" & _
  ",Coil_202,Coil_277"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_53,Coil_128" & _
  ",Coil_203,Coil_278"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_54,Coil_129" & _
  ",Coil_204,Coil_279"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_55,Coil_130" & _
  ",Coil_205,Coil_280"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_56,Coil_131" & _
  ",Coil_206,Coil_281"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_57,Coil_132" & _
  ",Coil_207,Coil_282"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_58,Coil_133" & _
  ",Coil_208,Coil_283"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_59,Coil_134" & _
  ",Coil_209,Coil_284"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_60,Coil_135" & _
  ",Coil_210,Coil_285"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_61,Coil_136" & _
  ",Coil_211,Coil_286"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_62,Coil_137" & _
  ",Coil_212,Coil_287"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_63,Coil_138" & _
  ",Coil_213,Coil_288"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_64,Coil_139" & _
  ",Coil_214,Coil_289"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_65,Coil_140" & _
  ",Coil_215,Coil_290"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_66,Coil_141" & _
  ",Coil_216,Coil_291"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_67,Coil_142" & _
  ",Coil_217,Coil_292"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_68,Coil_143" & _
  ",Coil_218,Coil_293"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_69,Coil_144" & _
  ",Coil_219,Coil_294"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_70,Coil_145" & _
  ",Coil_220,Coil_295"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_71,Coil_146" & _
  ",Coil_221,Coil_296"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_72,Coil_147" & _
  ",Coil_222,Coil_297"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_73,Coil_148" & _
  ",Coil_223,Coil_298"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_74,Coil_149" & _
  ",Coil_224,Coil_299"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/LapCoil", "Version:=", "12.0", "NoOfParameters:=", 21, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5431mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "5974mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", "Value:=", "0deg"), _
  Array("NAME:Pair", "Name:=", "Slots", "Value:=", "300"), Array("NAME:Pair", _
  "Name:=", "SlotType", "Value:=", "6"), Array("NAME:Pair", "Name:=", "Hs0", _
  "Value:=", "1mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "4.2mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "120mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "22.7mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "22.7mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", _
  "22.7mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", _
  "Name:=", "Layers", "Value:=", "2"), Array("NAME:Pair", "Name:=", _
  "CoilPitch", "Value:=", "9"), Array("NAME:Pair", "Name:=", "EndExt", _
  "Value:=", "150mm"), Array("NAME:Pair", "Name:=", "SpanExt", "Value:=", _
  "0mm"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", "10deg"), _
  Array("NAME:Pair", "Name:=", "LenRegion", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "InfoCoil", "Value:=", "3"))), Array(_
  "NAME:Attributes", "Name:=", "CoilRe", "Flags:=", "", "Color:=", _
  "(250 167 14)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "copper", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "CoilRe:CreateUserDefinedPart:1", _
  "LenRegion", "1180mm + 2*endRegion"
On Error Goto 0
oEditor.Rotate Array("NAME:Selections", "Selections:=", "CoilRe"), Array(_
  "NAME:RotateParameters", "CoordinateSystemID:=", -1, "RotateAxis:=", "Z", _
  "RotateAngle:=", "-10.799999999999999deg")
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "CoilRe"), _
  Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, _
  "CreateNewObjects:=", true, "WhichAxis:=", "Z", "AngleStr:=", "1.2deg", _
  "NumClones:=", "300"), Array("NAME:Options", "DuplicateBoundaries:=", _
  false)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "CoilRe"), Array(_
  "NAME:ChangedProps", Array("NAME:Name", "Value:=", "CoilRe_0"))))
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_0,CoilRe_75" & _
  ",CoilRe_150,CoilRe_225"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_1,CoilRe_76" & _
  ",CoilRe_151,CoilRe_226"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_2,CoilRe_77" & _
  ",CoilRe_152,CoilRe_227"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_3,CoilRe_78" & _
  ",CoilRe_153,CoilRe_228"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_4,CoilRe_79" & _
  ",CoilRe_154,CoilRe_229"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_5,CoilRe_80" & _
  ",CoilRe_155,CoilRe_230"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_6,CoilRe_81" & _
  ",CoilRe_156,CoilRe_231"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_7,CoilRe_82" & _
  ",CoilRe_157,CoilRe_232"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_8,CoilRe_83" & _
  ",CoilRe_158,CoilRe_233"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_9,CoilRe_84" & _
  ",CoilRe_159,CoilRe_234"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_10,CoilRe_85" & _
  ",CoilRe_160,CoilRe_235"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_11,CoilRe_86" & _
  ",CoilRe_161,CoilRe_236"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_12,CoilRe_87" & _
  ",CoilRe_162,CoilRe_237"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_13,CoilRe_88" & _
  ",CoilRe_163,CoilRe_238"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_14,CoilRe_89" & _
  ",CoilRe_164,CoilRe_239"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_15,CoilRe_90" & _
  ",CoilRe_165,CoilRe_240"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_16,CoilRe_91" & _
  ",CoilRe_166,CoilRe_241"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_17,CoilRe_92" & _
  ",CoilRe_167,CoilRe_242"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_18,CoilRe_93" & _
  ",CoilRe_168,CoilRe_243"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_19,CoilRe_94" & _
  ",CoilRe_169,CoilRe_244"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_20,CoilRe_95" & _
  ",CoilRe_170,CoilRe_245"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_21,CoilRe_96" & _
  ",CoilRe_171,CoilRe_246"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_22,CoilRe_97" & _
  ",CoilRe_172,CoilRe_247"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_23,CoilRe_98" & _
  ",CoilRe_173,CoilRe_248"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_24,CoilRe_99" & _
  ",CoilRe_174,CoilRe_249"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_25,CoilRe_100" & _
  ",CoilRe_175,CoilRe_250"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_26,CoilRe_101" & _
  ",CoilRe_176,CoilRe_251"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_27,CoilRe_102" & _
  ",CoilRe_177,CoilRe_252"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_28,CoilRe_103" & _
  ",CoilRe_178,CoilRe_253"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_29,CoilRe_104" & _
  ",CoilRe_179,CoilRe_254"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_30,CoilRe_105" & _
  ",CoilRe_180,CoilRe_255"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_31,CoilRe_106" & _
  ",CoilRe_181,CoilRe_256"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_32,CoilRe_107" & _
  ",CoilRe_182,CoilRe_257"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_33,CoilRe_108" & _
  ",CoilRe_183,CoilRe_258"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_34,CoilRe_109" & _
  ",CoilRe_184,CoilRe_259"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_35,CoilRe_110" & _
  ",CoilRe_185,CoilRe_260"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_36,CoilRe_111" & _
  ",CoilRe_186,CoilRe_261"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_37,CoilRe_112" & _
  ",CoilRe_187,CoilRe_262"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_38,CoilRe_113" & _
  ",CoilRe_188,CoilRe_263"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_39,CoilRe_114" & _
  ",CoilRe_189,CoilRe_264"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_40,CoilRe_115" & _
  ",CoilRe_190,CoilRe_265"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_41,CoilRe_116" & _
  ",CoilRe_191,CoilRe_266"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_42,CoilRe_117" & _
  ",CoilRe_192,CoilRe_267"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_43,CoilRe_118" & _
  ",CoilRe_193,CoilRe_268"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_44,CoilRe_119" & _
  ",CoilRe_194,CoilRe_269"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_45,CoilRe_120" & _
  ",CoilRe_195,CoilRe_270"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_46,CoilRe_121" & _
  ",CoilRe_196,CoilRe_271"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_47,CoilRe_122" & _
  ",CoilRe_197,CoilRe_272"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_48,CoilRe_123" & _
  ",CoilRe_198,CoilRe_273"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_49,CoilRe_124" & _
  ",CoilRe_199,CoilRe_274"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_50,CoilRe_125" & _
  ",CoilRe_200,CoilRe_275"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_51,CoilRe_126" & _
  ",CoilRe_201,CoilRe_276"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_52,CoilRe_127" & _
  ",CoilRe_202,CoilRe_277"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_53,CoilRe_128" & _
  ",CoilRe_203,CoilRe_278"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_54,CoilRe_129" & _
  ",CoilRe_204,CoilRe_279"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_55,CoilRe_130" & _
  ",CoilRe_205,CoilRe_280"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_56,CoilRe_131" & _
  ",CoilRe_206,CoilRe_281"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_57,CoilRe_132" & _
  ",CoilRe_207,CoilRe_282"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_58,CoilRe_133" & _
  ",CoilRe_208,CoilRe_283"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_59,CoilRe_134" & _
  ",CoilRe_209,CoilRe_284"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_60,CoilRe_135" & _
  ",CoilRe_210,CoilRe_285"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_61,CoilRe_136" & _
  ",CoilRe_211,CoilRe_286"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_62,CoilRe_137" & _
  ",CoilRe_212,CoilRe_287"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_63,CoilRe_138" & _
  ",CoilRe_213,CoilRe_288"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_64,CoilRe_139" & _
  ",CoilRe_214,CoilRe_289"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_65,CoilRe_140" & _
  ",CoilRe_215,CoilRe_290"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_66,CoilRe_141" & _
  ",CoilRe_216,CoilRe_291"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_67,CoilRe_142" & _
  ",CoilRe_217,CoilRe_292"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_68,CoilRe_143" & _
  ",CoilRe_218,CoilRe_293"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_69,CoilRe_144" & _
  ",CoilRe_219,CoilRe_294"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_70,CoilRe_145" & _
  ",CoilRe_220,CoilRe_295"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_71,CoilRe_146" & _
  ",CoilRe_221,CoilRe_296"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_72,CoilRe_147" & _
  ",CoilRe_222,CoilRe_297"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_73,CoilRe_148" & _
  ",CoilRe_223,CoilRe_298"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_74,CoilRe_149" & _
  ",CoilRe_224,CoilRe_299"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignWindingGroup Array("NAME:PhaseA", "Type:=", "Voltage", _
  "IsSolid:=", false, "Current:=", "0A", "Voltage:=", _
  "11267.7*sin(2*pi*50*time+delta)", "Resistance:=", "R1", "Inductance:=", _
  "Le1", "ParallelBranchesNum:=", "1")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignWindingGroup Array("NAME:PhaseB", "Type:=", "Voltage", _
  "IsSolid:=", false, "Current:=", "0A", "Voltage:=", _
  "11267.7*sin(2*pi*50*time+delta-2*pi/3)", "Resistance:=", "R1", _
  "Inductance:=", "Le1", "ParallelBranchesNum:=", "1")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignWindingGroup Array("NAME:PhaseC", "Type:=", "Voltage", _
  "IsSolid:=", false, "Current:=", "0A", "Voltage:=", _
  "11267.7*sin(2*pi*50*time+delta-4*pi/3)", "Resistance:=", "R1", _
  "Inductance:=", "Le1", "ParallelBranchesNum:=", "1")
oModule.AssignCoil Array("NAME:PhA_0", "Objects:=", Array("Coil_0"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_0", "Objects:=", Array("CoilRe_9"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_1", "Objects:=", Array("Coil_1"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_1", "Objects:=", Array("CoilRe_10"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_2", "Objects:=", Array("Coil_2"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_2", "Objects:=", Array("CoilRe_11"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_3", "Objects:=", Array("Coil_3"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_3", "Objects:=", Array("CoilRe_12"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhCRe_4", "Objects:=", Array("Coil_4"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_4", "Objects:=", Array("CoilRe_13"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_5", "Objects:=", Array("Coil_5"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_5", "Objects:=", Array("CoilRe_14"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_6", "Objects:=", Array("Coil_6"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_6", "Objects:=", Array("CoilRe_15"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhB_7", "Objects:=", Array("Coil_7"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_7", "Objects:=", Array("CoilRe_16"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_8", "Objects:=", Array("Coil_8"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_8", "Objects:=", Array("CoilRe_17"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_9", "Objects:=", Array("Coil_9"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_9", "Objects:=", Array("CoilRe_18"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhARe_10", "Objects:=", Array("Coil_10"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_10", "Objects:=", Array("CoilRe_19"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_11", "Objects:=", Array("Coil_11"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_11", "Objects:=", Array("CoilRe_20"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_12", "Objects:=", Array("Coil_12"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_12", "Objects:=", Array("CoilRe_21"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhC_13", "Objects:=", Array("Coil_13"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_13", "Objects:=", Array("CoilRe_22"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_14", "Objects:=", Array("Coil_14"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_14", "Objects:=", Array("CoilRe_23"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_15", "Objects:=", Array("Coil_15"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_15", "Objects:=", Array("CoilRe_24"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhBRe_16", "Objects:=", Array("Coil_16"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_16", "Objects:=", Array("CoilRe_25"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_17", "Objects:=", Array("Coil_17"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_17", "Objects:=", Array("CoilRe_26"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_18", "Objects:=", Array("Coil_18"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_18", "Objects:=", Array("CoilRe_27"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhA_19", "Objects:=", Array("Coil_19"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_19", "Objects:=", Array("CoilRe_28"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_20", "Objects:=", Array("Coil_20"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_20", "Objects:=", Array("CoilRe_29"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_21", "Objects:=", Array("Coil_21"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_21", "Objects:=", Array("CoilRe_30"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhCRe_22", "Objects:=", Array("Coil_22"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_22", "Objects:=", Array("CoilRe_31"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_23", "Objects:=", Array("Coil_23"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_23", "Objects:=", Array("CoilRe_32"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_24", "Objects:=", Array("Coil_24"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_24", "Objects:=", Array("CoilRe_33"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhB_25", "Objects:=", Array("Coil_25"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_25", "Objects:=", Array("CoilRe_34"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_26", "Objects:=", Array("Coil_26"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_26", "Objects:=", Array("CoilRe_35"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_27", "Objects:=", Array("Coil_27"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_27", "Objects:=", Array("CoilRe_36"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_28", "Objects:=", Array("Coil_28"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_28", "Objects:=", Array("CoilRe_37"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhARe_29", "Objects:=", Array("Coil_29"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_29", "Objects:=", Array("CoilRe_38"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_30", "Objects:=", Array("Coil_30"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_30", "Objects:=", Array("CoilRe_39"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_31", "Objects:=", Array("Coil_31"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_31", "Objects:=", Array("CoilRe_40"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhC_32", "Objects:=", Array("Coil_32"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_32", "Objects:=", Array("CoilRe_41"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_33", "Objects:=", Array("Coil_33"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_33", "Objects:=", Array("CoilRe_42"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_34", "Objects:=", Array("Coil_34"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_34", "Objects:=", Array("CoilRe_43"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhBRe_35", "Objects:=", Array("Coil_35"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_35", "Objects:=", Array("CoilRe_44"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_36", "Objects:=", Array("Coil_36"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_36", "Objects:=", Array("CoilRe_45"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_37", "Objects:=", Array("Coil_37"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_37", "Objects:=", Array("CoilRe_46"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhA_38", "Objects:=", Array("Coil_38"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_38", "Objects:=", Array("CoilRe_47"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_39", "Objects:=", Array("Coil_39"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_39", "Objects:=", Array("CoilRe_48"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_40", "Objects:=", Array("Coil_40"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_40", "Objects:=", Array("CoilRe_49"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhCRe_41", "Objects:=", Array("Coil_41"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_41", "Objects:=", Array("CoilRe_50"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_42", "Objects:=", Array("Coil_42"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_42", "Objects:=", Array("CoilRe_51"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_43", "Objects:=", Array("Coil_43"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_43", "Objects:=", Array("CoilRe_52"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhB_44", "Objects:=", Array("Coil_44"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_44", "Objects:=", Array("CoilRe_53"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_45", "Objects:=", Array("Coil_45"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_45", "Objects:=", Array("CoilRe_54"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_46", "Objects:=", Array("Coil_46"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_46", "Objects:=", Array("CoilRe_55"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhARe_47", "Objects:=", Array("Coil_47"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_47", "Objects:=", Array("CoilRe_56"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_48", "Objects:=", Array("Coil_48"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_48", "Objects:=", Array("CoilRe_57"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_49", "Objects:=", Array("Coil_49"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_49", "Objects:=", Array("CoilRe_58"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhC_50", "Objects:=", Array("Coil_50"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_50", "Objects:=", Array("CoilRe_59"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_51", "Objects:=", Array("Coil_51"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_51", "Objects:=", Array("CoilRe_60"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_52", "Objects:=", Array("Coil_52"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_52", "Objects:=", Array("CoilRe_61"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_53", "Objects:=", Array("Coil_53"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_53", "Objects:=", Array("CoilRe_62"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhBRe_54", "Objects:=", Array("Coil_54"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_54", "Objects:=", Array("CoilRe_63"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_55", "Objects:=", Array("Coil_55"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_55", "Objects:=", Array("CoilRe_64"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_56", "Objects:=", Array("Coil_56"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_56", "Objects:=", Array("CoilRe_65"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhA_57", "Objects:=", Array("Coil_57"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_57", "Objects:=", Array("CoilRe_66"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_58", "Objects:=", Array("Coil_58"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_58", "Objects:=", Array("CoilRe_67"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_59", "Objects:=", Array("Coil_59"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_59", "Objects:=", Array("CoilRe_68"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhCRe_60", "Objects:=", Array("Coil_60"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_60", "Objects:=", Array("CoilRe_69"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_61", "Objects:=", Array("Coil_61"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_61", "Objects:=", Array("CoilRe_70"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_62", "Objects:=", Array("Coil_62"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_62", "Objects:=", Array("CoilRe_71"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhB_63", "Objects:=", Array("Coil_63"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_63", "Objects:=", Array("CoilRe_72"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_64", "Objects:=", Array("Coil_64"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_64", "Objects:=", Array("CoilRe_73"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_65", "Objects:=", Array("Coil_65"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_65", "Objects:=", Array("CoilRe_74"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhARe_66", "Objects:=", Array("Coil_66"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_67", "Objects:=", Array("Coil_67"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_68", "Objects:=", Array("Coil_68"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhC_69", "Objects:=", Array("Coil_69"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_70", "Objects:=", Array("Coil_70"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_71", "Objects:=", Array("Coil_71"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhBRe_72", "Objects:=", Array("Coil_72"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_73", "Objects:=", Array("Coil_73"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_74", "Objects:=", Array("Coil_74"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhA_291", "Objects:=", Array("CoilRe_0"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_292", "Objects:=", Array("CoilRe_1"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_293", "Objects:=", Array("CoilRe_2"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhCRe_294", "Objects:=", Array("CoilRe_3"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_295", "Objects:=", Array("CoilRe_4"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_296", "Objects:=", Array("CoilRe_5"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhB_297", "Objects:=", Array("CoilRe_6"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_298", "Objects:=", Array("CoilRe_7"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_299", "Objects:=", Array("CoilRe_8"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "5398mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "1800mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "2"), _
  Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs01", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", "10mm"), _
  Array("NAME:Pair", "Name:=", "Rs", "Value:=", "5mm"), Array("NAME:Pair", _
  "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", _
  "SlotPitch", "Value:=", "1deg"), Array("NAME:Pair", "Name:=", "CenterPitch", _
  "Value:=", "2deg"), Array("NAME:Pair", "Name:=", "Poles", "Value:=", "32"), _
  Array("NAME:Pair", "Name:=", "WidthShoe", "Value:=", "355mm"), Array(_
  "NAME:Pair", "Name:=", "HeightShoe", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "WidthBody", "Value:=", "240mm"), Array("NAME:Pair", "Name:=", _
  "HeightBody", "Value:=", "300mm"), Array("NAME:Pair", "Name:=", "AirGap2", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Offset", "Value:=", "10mm"), _
  Array("NAME:Pair", "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "CoilEndExt", "Value:=", "41.5mm"), Array("NAME:Pair", "Name:=", _
  "EndRingType", "Value:=", "1"), Array("NAME:Pair", "Name:=", "BarEndExt", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "RingLength", "Value:=", _
  "40mm"), Array("NAME:Pair", "Name:=", "RingHeight", "Value:=", "21mm"), _
  Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", "15deg"), Array(_
  "NAME:Pair", "Name:=", "LenRegion", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "InfoCore", "Value:=", "0"))), Array("NAME:Attributes", "Name:=", _
  "Rotor", "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=", 0, _
  "PartCoordinateSystem:=", "Global", "MaterialName:=", "M22_24G_2DSF0.999", _
  "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Rotor:CreateUserDefinedPart:1", _
  "LenRegion", "1180mm + 2*endRegion"
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "5398mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "1800mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "2"), _
  Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs01", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", "10mm"), _
  Array("NAME:Pair", "Name:=", "Rs", "Value:=", "5mm"), Array("NAME:Pair", _
  "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", _
  "SlotPitch", "Value:=", "1deg"), Array("NAME:Pair", "Name:=", "CenterPitch", _
  "Value:=", "2deg"), Array("NAME:Pair", "Name:=", "Poles", "Value:=", "32"), _
  Array("NAME:Pair", "Name:=", "WidthShoe", "Value:=", "355mm"), Array(_
  "NAME:Pair", "Name:=", "HeightShoe", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "WidthBody", "Value:=", "240mm"), Array("NAME:Pair", "Name:=", _
  "HeightBody", "Value:=", "300mm"), Array("NAME:Pair", "Name:=", "AirGap2", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Offset", "Value:=", "10mm"), _
  Array("NAME:Pair", "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "CoilEndExt", "Value:=", "41.5mm"), Array("NAME:Pair", "Name:=", _
  "EndRingType", "Value:=", "1"), Array("NAME:Pair", "Name:=", "BarEndExt", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "RingLength", "Value:=", _
  "40mm"), Array("NAME:Pair", "Name:=", "RingHeight", "Value:=", "21mm"), _
  Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", "15deg"), Array(_
  "NAME:Pair", "Name:=", "LenRegion", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "InfoCore", "Value:=", "4"))), Array("NAME:Attributes", "Name:=", _
  "Field", "Flags:=", "", "Color:=", "(255 128 192)", "Transparency:=", 0, _
  "PartCoordinateSystem:=", "Global", "MaterialName:=", "copper", _
  "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Field:CreateUserDefinedPart:1", _
  "LenRegion", "1180mm + 2*endRegion"
On Error Goto 0
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "Field"), _
  Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, _
  "CreateNewObjects:=", true, "WhichAxis:=", "Z", "AngleStr:=", "11.25deg", _
  "NumClones:=", "32"), Array("NAME:Options", "DuplicateBoundaries:=", false)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "Field"), Array(_
  "NAME:ChangedProps", Array("NAME:Name", "Value:=", "Field_0"))))
oEditor.Unite Array("NAME:Selections", "Selections:=", "Field_0,Field_8" & _
  ",Field_16,Field_24"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Field_1,Field_9" & _
  ",Field_17,Field_25"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Field_2,Field_10" & _
  ",Field_18,Field_26"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Field_3,Field_11" & _
  ",Field_19,Field_27"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Field_4,Field_12" & _
  ",Field_20,Field_28"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Field_5,Field_13" & _
  ",Field_21,Field_29"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Field_6,Field_14" & _
  ",Field_22,Field_30"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Field_7,Field_15" & _
  ",Field_23,Field_31"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "5398mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "1800mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "2"), _
  Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs01", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", "10mm"), _
  Array("NAME:Pair", "Name:=", "Rs", "Value:=", "5mm"), Array("NAME:Pair", _
  "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", _
  "SlotPitch", "Value:=", "1deg"), Array("NAME:Pair", "Name:=", "CenterPitch", _
  "Value:=", "2deg"), Array("NAME:Pair", "Name:=", "Poles", "Value:=", "32"), _
  Array("NAME:Pair", "Name:=", "WidthShoe", "Value:=", "355mm"), Array(_
  "NAME:Pair", "Name:=", "HeightShoe", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "WidthBody", "Value:=", "240mm"), Array("NAME:Pair", "Name:=", _
  "HeightBody", "Value:=", "300mm"), Array("NAME:Pair", "Name:=", "AirGap2", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Offset", "Value:=", "10mm"), _
  Array("NAME:Pair", "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "CoilEndExt", "Value:=", "41.5mm"), Array("NAME:Pair", "Name:=", _
  "EndRingType", "Value:=", "1"), Array("NAME:Pair", "Name:=", "BarEndExt", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "RingLength", "Value:=", _
  "40mm"), Array("NAME:Pair", "Name:=", "RingHeight", "Value:=", "21mm"), _
  Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", "15deg"), Array(_
  "NAME:Pair", "Name:=", "LenRegion", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "InfoCore", "Value:=", "5"))), Array("NAME:Attributes", "Name:=", _
  "FieldRe", "Flags:=", "", "Color:=", "(255 128 192)", "Transparency:=", 0, _
  "PartCoordinateSystem:=", "Global", "MaterialName:=", "copper", _
  "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "FieldRe:CreateUserDefinedPart:1", _
  "LenRegion", "1180mm + 2*endRegion"
On Error Goto 0
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "FieldRe"), _
  Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, _
  "CreateNewObjects:=", true, "WhichAxis:=", "Z", "AngleStr:=", "11.25deg", _
  "NumClones:=", "32"), Array("NAME:Options", "DuplicateBoundaries:=", false)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "FieldRe"), Array(_
  "NAME:ChangedProps", Array("NAME:Name", "Value:=", "FieldRe_0"))))
oEditor.Unite Array("NAME:Selections", "Selections:=", "FieldRe_0,FieldRe_8" & _
  ",FieldRe_16,FieldRe_24"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "FieldRe_1,FieldRe_9" & _
  ",FieldRe_17,FieldRe_25"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "FieldRe_2,FieldRe_10" & _
  ",FieldRe_18,FieldRe_26"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "FieldRe_3,FieldRe_11" & _
  ",FieldRe_19,FieldRe_27"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "FieldRe_4,FieldRe_12" & _
  ",FieldRe_20,FieldRe_28"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "FieldRe_5,FieldRe_13" & _
  ",FieldRe_21,FieldRe_29"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "FieldRe_6,FieldRe_14" & _
  ",FieldRe_22,FieldRe_30"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "FieldRe_7,FieldRe_15" & _
  ",FieldRe_23,FieldRe_31"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignWindingGroup Array("NAME:Field", "Type:=", "Current", _
  "IsSolid:=", false, "Current:=", "203.396A", "Resistance:=", "0ohm", _
  "Inductance:=", "0H", "ParallelBranchesNum:=", "1")
oModule.AssignCoil Array("NAME:Field_0", "Objects:=", Array("Field_0"), _
  "Conductor number:=", 150, "PolarityType:=", "Positive", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:FieldRe_0", "Objects:=", Array("FieldRe_0"), _
  "Conductor number:=", 150, "PolarityType:=", "Negative", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:FieldRe_1", "Objects:=", Array("Field_1"), _
  "Conductor number:=", 150, "PolarityType:=", "Negative", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:Field_1", "Objects:=", Array("FieldRe_1"), _
  "Conductor number:=", 150, "PolarityType:=", "Positive", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:Field_2", "Objects:=", Array("Field_2"), _
  "Conductor number:=", 150, "PolarityType:=", "Positive", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:FieldRe_2", "Objects:=", Array("FieldRe_2"), _
  "Conductor number:=", 150, "PolarityType:=", "Negative", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:FieldRe_3", "Objects:=", Array("Field_3"), _
  "Conductor number:=", 150, "PolarityType:=", "Negative", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:Field_3", "Objects:=", Array("FieldRe_3"), _
  "Conductor number:=", 150, "PolarityType:=", "Positive", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:Field_4", "Objects:=", Array("Field_4"), _
  "Conductor number:=", 150, "PolarityType:=", "Positive", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:FieldRe_4", "Objects:=", Array("FieldRe_4"), _
  "Conductor number:=", 150, "PolarityType:=", "Negative", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:FieldRe_5", "Objects:=", Array("Field_5"), _
  "Conductor number:=", 150, "PolarityType:=", "Negative", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:Field_5", "Objects:=", Array("FieldRe_5"), _
  "Conductor number:=", 150, "PolarityType:=", "Positive", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:Field_6", "Objects:=", Array("Field_6"), _
  "Conductor number:=", 150, "PolarityType:=", "Positive", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:FieldRe_6", "Objects:=", Array("FieldRe_6"), _
  "Conductor number:=", 150, "PolarityType:=", "Negative", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:FieldRe_7", "Objects:=", Array("Field_7"), _
  "Conductor number:=", 150, "PolarityType:=", "Negative", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:Field_7", "Objects:=", Array("FieldRe_7"), _
  "Conductor number:=", 150, "PolarityType:=", "Positive", "Winding:=", _
  "Field")
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "5398mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "1800mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "2"), _
  Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs01", _
  "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", "10mm"), _
  Array("NAME:Pair", "Name:=", "Rs", "Value:=", "5mm"), Array("NAME:Pair", _
  "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", _
  "SlotPitch", "Value:=", "1deg"), Array("NAME:Pair", "Name:=", "CenterPitch", _
  "Value:=", "2deg"), Array("NAME:Pair", "Name:=", "Poles", "Value:=", "32"), _
  Array("NAME:Pair", "Name:=", "WidthShoe", "Value:=", "355mm"), Array(_
  "NAME:Pair", "Name:=", "HeightShoe", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "WidthBody", "Value:=", "240mm"), Array("NAME:Pair", "Name:=", _
  "HeightBody", "Value:=", "300mm"), Array("NAME:Pair", "Name:=", "AirGap2", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Offset", "Value:=", "10mm"), _
  Array("NAME:Pair", "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "CoilEndExt", "Value:=", "41.5mm"), Array("NAME:Pair", "Name:=", _
  "EndRingType", "Value:=", "1"), Array("NAME:Pair", "Name:=", "BarEndExt", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "RingLength", "Value:=", _
  "40mm"), Array("NAME:Pair", "Name:=", "RingHeight", "Value:=", "21mm"), _
  Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", "15deg"), Array(_
  "NAME:Pair", "Name:=", "LenRegion", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "InfoCore", "Value:=", "3"))), Array("NAME:Attributes", "Name:=", _
  "Bar", "Flags:=", "", "Color:=", "(119 198 206)", "Transparency:=", 0, _
  "PartCoordinateSystem:=", "Global", "MaterialName:=", "copper_100C", _
  "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Bar:CreateUserDefinedPart:1", _
  "LenRegion", "1180mm + 2*endRegion"
On Error Goto 0
oEditor.Split Array("NAME:Selections", "Selections:=", "Bar", _
  "NewPartsModelFlag:=", "Model"), Array("NAME:SplitToParameters", _
  "CoordinateSystemID:=", -1, "SplitPlane:=", "ZX", "WhichSide:=", _
  "PositiveOnly", "SplitCrossingObjectsOnly:=", true)
oEditor.Rotate Array("NAME:Selections", "Selections:=", "Bar"), Array(_
  "NAME:RotateParameters", "CoordinateSystemID:=", -1, "RotateAxis:=", "Z", _
  "RotateAngle:=", "168.75deg")
oEditor.Split Array("NAME:Selections", "Selections:=", "Bar", _
  "NewPartsModelFlag:=", "Model"), Array("NAME:SplitToParameters", _
  "CoordinateSystemID:=", -1, "SplitPlane:=", "ZX", "WhichSide:=", _
  "PositiveOnly", "SplitCrossingObjectsOnly:=", true)
oEditor.Rotate Array("NAME:Selections", "Selections:=", "Bar"), Array(_
  "NAME:RotateParameters", "CoordinateSystemID:=", -1, "RotateAxis:=", "Z", _
  "RotateAngle:=", "-168.75deg")
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "Bar"), _
  Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, _
  "CreateNewObjects:=", true, "WhichAxis:=", "Z", "AngleStr:=", "11.25deg", _
  "NumClones:=", "8"), Array("NAME:Options", "DuplicateBoundaries:=", false)
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Bar", _
  "Tool Parts:=", "Tool"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", true)
oEditor.SeparateBody  Array("NAME:Selections", "Selections:=", "Bar")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignEndConnection Array("NAME:EndConnection1", "Objects:=", Array(_
  "Bar", "Bar_Separate1"), "ResistanceValue:=", "0ohm", "InductanceValue:=", _
  "0H")
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Bar_1", _
  "Tool Parts:=", "Tool"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", true)
oEditor.SeparateBody  Array("NAME:Selections", "Selections:=", "Bar_1")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignEndConnection Array("NAME:EndConnection2", "Objects:=", Array(_
  "Bar_1", "Bar_1_Separate1"), "ResistanceValue:=", "0ohm", _
  "InductanceValue:=", "0H")
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Bar_2", _
  "Tool Parts:=", "Tool"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", true)
oEditor.SeparateBody  Array("NAME:Selections", "Selections:=", "Bar_2")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignEndConnection Array("NAME:EndConnection3", "Objects:=", Array(_
  "Bar_2", "Bar_2_Separate1"), "ResistanceValue:=", "0ohm", _
  "InductanceValue:=", "0H")
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Bar_3", _
  "Tool Parts:=", "Tool"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", true)
oEditor.SeparateBody  Array("NAME:Selections", "Selections:=", "Bar_3")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignEndConnection Array("NAME:EndConnection4", "Objects:=", Array(_
  "Bar_3", "Bar_3_Separate1"), "ResistanceValue:=", "0ohm", _
  "InductanceValue:=", "0H")
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Bar_4", _
  "Tool Parts:=", "Tool"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", true)
oEditor.SeparateBody  Array("NAME:Selections", "Selections:=", "Bar_4")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignEndConnection Array("NAME:EndConnection5", "Objects:=", Array(_
  "Bar_4", "Bar_4_Separate1"), "ResistanceValue:=", "0ohm", _
  "InductanceValue:=", "0H")
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Bar_5", _
  "Tool Parts:=", "Tool"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", true)
oEditor.SeparateBody  Array("NAME:Selections", "Selections:=", "Bar_5")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignEndConnection Array("NAME:EndConnection6", "Objects:=", Array(_
  "Bar_5", "Bar_5_Separate1"), "ResistanceValue:=", "0ohm", _
  "InductanceValue:=", "0H")
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Bar_6", _
  "Tool Parts:=", "Tool"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", true)
oEditor.SeparateBody  Array("NAME:Selections", "Selections:=", "Bar_6")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignEndConnection Array("NAME:EndConnection7", "Objects:=", Array(_
  "Bar_6", "Bar_6_Separate1"), "ResistanceValue:=", "0ohm", _
  "InductanceValue:=", "0H")
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Bar_7", _
  "Tool Parts:=", "Tool"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", true)
oEditor.SeparateBody  Array("NAME:Selections", "Selections:=", "Bar_7")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignEndConnection Array("NAME:EndConnection8", "Objects:=", Array(_
  "Bar_7", "Bar_7_Separate1"), "ResistanceValue:=", "0ohm", _
  "InductanceValue:=", "0H")
Set oModule = oDesign.GetModule("MeshSetup")
oModule.AssignTrueSurfOp Array("NAME:SurfApprox_Bar", "Objects:=", Array("Bar", _
  "Bar_Separate1", "Bar_1", "Bar_1_Separate1", "Bar_2", "Bar_2_Separate1", _
  "Bar_3", "Bar_3_Separate1", "Bar_4", "Bar_4_Separate1", "Bar_5", _
  "Bar_5_Separate1", "Bar_6", "Bar_6_Separate1", "Bar_7", "Bar_7_Separate1"), _
  "SurfDevChoice:=", 2, "SurfDev:=", "2.699mm", "NormalDevChoice:=", 2, _
  "NormalDev:=", "15deg", "AspectRatioChoice:=", 1)
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "5398mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "1800mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "2"), _
  Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs01", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", "10mm"), _
  Array("NAME:Pair", "Name:=", "Rs", "Value:=", "5mm"), Array("NAME:Pair", _
  "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", _
  "SlotPitch", "Value:=", "1deg"), Array("NAME:Pair", "Name:=", "CenterPitch", _
  "Value:=", "2deg"), Array("NAME:Pair", "Name:=", "Poles", "Value:=", "32"), _
  Array("NAME:Pair", "Name:=", "WidthShoe", "Value:=", "355mm"), Array(_
  "NAME:Pair", "Name:=", "HeightShoe", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "WidthBody", "Value:=", "240mm"), Array("NAME:Pair", "Name:=", _
  "HeightBody", "Value:=", "300mm"), Array("NAME:Pair", "Name:=", "AirGap2", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Offset", "Value:=", "10mm"), _
  Array("NAME:Pair", "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "CoilEndExt", "Value:=", "41.5mm"), Array("NAME:Pair", "Name:=", _
  "EndRingType", "Value:=", "1"), Array("NAME:Pair", "Name:=", "BarEndExt", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "RingLength", "Value:=", _
  "40mm"), Array("NAME:Pair", "Name:=", "RingHeight", "Value:=", "21mm"), _
  Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", "15deg"), Array(_
  "NAME:Pair", "Name:=", "LenRegion", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "InfoCore", "Value:=", "100"))), Array("NAME:Attributes", _
  "Name:=", "InnerRegion", "Flags:=", "", "Color:=", "(0 255 255)", _
  "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "MaterialName:=", _
  "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "InnerRegion:CreateUserDefinedPart:1", "LenRegion", "1180mm + 2*endRegion"
On Error Goto 0
Set oModule = oDesign.GetModule("MeshSetup")
oModule.AssignTrueSurfOp Array("NAME:SurfApprox_Main", "Objects:=", Array(_
  "Stator", "Rotor", "Band", "OuterRegion", "InnerRegion", "Shaft"), _
  "SurfDevChoice:=", 2, "SurfDev:=", "2.987mm", "NormalDevChoice:=", 2, _
  "NormalDev:=", "15deg", "AspectRatioChoice:=", 1)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "Band", _
  "OuterRegion", "InnerRegion"), Array("NAME:ChangedProps", Array(_
  "NAME:Transparent", "Value:=", 0.75))))
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Band,InnerRegion" & _
  ",Shaft,Stator,Coil_0,Coil_1,Coil_2,Coil_3,Coil_4,Coil_5,Coil_6,Coil_7" & _
  ",Coil_8,Coil_9,Coil_10,Coil_11,Coil_12,Coil_13,Coil_14,Coil_15,Coil_16" & _
  ",Coil_17,Coil_18,Coil_19,Coil_20,Coil_21,Coil_22,Coil_23,Coil_24" & _
  ",Coil_25,Coil_26,Coil_27,Coil_28,Coil_29,Coil_30,Coil_31,Coil_32" & _
  ",Coil_33,Coil_34,Coil_35,Coil_36,Coil_37,Coil_38,Coil_39,Coil_40" & _
  ",Coil_41,Coil_42,Coil_43,Coil_44,Coil_45,Coil_46,Coil_47,Coil_48" & _
  ",Coil_49,Coil_50,Coil_51,Coil_52,Coil_53,Coil_54,Coil_55,Coil_56" & _
  ",Coil_57,Coil_58,Coil_59,Coil_60,Coil_61,Coil_62,Coil_63,Coil_64" & _
  ",Coil_65,Coil_66,Coil_67,Coil_68,Coil_69,Coil_70,Coil_71,Coil_72" & _
  ",Coil_73,Coil_74,CoilRe_0,CoilRe_1,CoilRe_2,CoilRe_3,CoilRe_4,CoilRe_5" & _
  ",CoilRe_6,CoilRe_7,CoilRe_8,CoilRe_9,CoilRe_10,CoilRe_11,CoilRe_12" & _
  ",CoilRe_13,CoilRe_14,CoilRe_15,CoilRe_16,CoilRe_17,CoilRe_18,CoilRe_19" & _
  ",CoilRe_20,CoilRe_21,CoilRe_22,CoilRe_23,CoilRe_24,CoilRe_25,CoilRe_26" & _
  ",CoilRe_27,CoilRe_28,CoilRe_29,CoilRe_30,CoilRe_31,CoilRe_32,CoilRe_33" & _
  ",CoilRe_34,CoilRe_35,CoilRe_36,CoilRe_37,CoilRe_38,CoilRe_39,CoilRe_40" & _
  ",CoilRe_41,CoilRe_42,CoilRe_43,CoilRe_44,CoilRe_45,CoilRe_46,CoilRe_47" & _
  ",CoilRe_48,CoilRe_49,CoilRe_50,CoilRe_51,CoilRe_52,CoilRe_53,CoilRe_54" & _
  ",CoilRe_55,CoilRe_56,CoilRe_57,CoilRe_58,CoilRe_59,CoilRe_60,CoilRe_61" & _
  ",CoilRe_62,CoilRe_63,CoilRe_64,CoilRe_65,CoilRe_66,CoilRe_67,CoilRe_68" & _
  ",CoilRe_69,CoilRe_70,CoilRe_71,CoilRe_72,CoilRe_73,CoilRe_74,Rotor" & _
  ",Field_0,Field_1,Field_2,Field_3,Field_4,Field_5,Field_6,Field_7" & _
  ",FieldRe_0,FieldRe_1,FieldRe_2,FieldRe_3,FieldRe_4,FieldRe_5,FieldRe_6" & _
  ",FieldRe_7", "Tool Parts:=", "Tool"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.FitAll
Set oModule = oDesign.GetModule("ModelSetup")
oModule.AssignBand Array("NAME:MotionSetup1", "Move Type:=", "Rotate", _
  "Coordinate System:=", "Global", "Axis:=", "Z", "Is Positive:=", true, _
  "InitPos:=", "13.425deg", "HasRotateLimit:=", false, "NonCylindrical:=", _
  false, "Consider Mechanical Transient:=", true, "Angular Velocity:=", _
  "187.5rpm", "Moment of Inertia:=", 715186, "Damping:=", 1465.24, _
  "Load Torque:=", _
  "if(speed<19.635, -102654*speed, -3.95763e+07/speed) - 102654*(speed-19.635)*10", _
  "Objects:=", Array("Band"))
oModule.EditMotionSetup "MotionSetup1", Array("NAME:Data", _
  "Consider Mechanical Transient:=", false)
Set oModule = oDesign.GetModule("AnalysisSetup")
oModule.InsertSetup "Transient", Array("NAME:Setup1", "FastReachSteadyState:=", _
  true, "AutoDetectSteadyState:=", true, "FrequencyOfAddedVoltageSource:=", _
  "50Hz", "StopTime:=", "0.2s", "TimeStep:=", "0.0002s")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.SetMinimumTimeStep "0.002ms" 
Set oModule = oDesign.GetModule("ReportSetup")
oModule.CreateReport "Torque", "Transient", "XY Plot", "Setup1 : Transient", _
  Array(), Array("Time:=", Array("All")), Array("X Component:=", "Time", _
  "Y Component:=", Array("Moving1.Torque")), Array()
oModule.CreateReport "Currents", "Transient", "XY Plot", "Setup1 : Transient", _
  Array(), Array("Time:=", Array("All")), Array("X Component:=", "Time", _
  "Y Component:=", Array("Current(PhaseA)", "Current(PhaseB)", _
  "Current(PhaseC)")), Array()
oModule.CreateReport "Induced Voltages", "Transient", "XY Plot", _
  "Setup1 : Transient", Array(), Array("Time:=", Array("All")), Array(_
  "X Component:=", "Time", "Y Component:=", Array("InducedVoltage(PhaseA)", _
  "InducedVoltage(PhaseB)", "InducedVoltage(PhaseC)")), Array()
oModule.CreateReport "Flux Linkages", "Transient", "XY Plot", _
  "Setup1 : Transient", Array(), Array("Time:=", Array("All")), Array(_
  "X Component:=", "Time", "Y Component:=", Array("FluxLinkage(PhaseA)", _
  "FluxLinkage(PhaseB)", "FluxLinkage(PhaseC)")), Array()
oModule.CreateReport "Voltages", "Transient", "XY Plot", "Setup1 : Transient", _
  Array(), Array("Time:=", Array("All")), Array("X Component:=", "Time", _
  "Y Component:=", Array("InputVoltage(PhaseA)", "InputVoltage(PhaseB)", _
  "InputVoltage(PhaseC)")), Array()
oModule.CreateReport "Powers", "Transient", "XY Plot", "Setup1 : Transient", _
  Array(), Array("Time:=", Array("All")), Array("X Component:=", "Time", _
  "Y Component:=", Array(_
  "InputVoltage(PhaseA)*Current(PhaseA) + InputVoltage(PhaseB)*Current(PhaseB) + InputVoltage(PhaseC)*Current(PhaseC)", _
  "Moving1.Speed*Moving1.Torque")), Array()
oModule.ChangeProperty Array("NAME:AllTabs", Array("NAME:Trace", Array(_
  "NAME:PropServers", _
  "Powers:InputVoltage(PhaseA)*Current(PhaseA) + InputVoltage(PhaseB)*Current(PhaseB) + InputVoltage(PhaseC)*Current(PhaseC)"), _
  Array("NAME:ChangedProps", Array("NAME:Specify Name", "Value:=", true))))
oModule.ChangeProperty Array("NAME:AllTabs", Array("NAME:Trace", Array(_
  "NAME:PropServers", _
  "Powers:InputVoltage(PhaseA)*Current(PhaseA) + InputVoltage(PhaseB)*Current(PhaseB) + InputVoltage(PhaseC)*Current(PhaseC)"), _
  Array("NAME:ChangedProps", Array("NAME:Name", "Value:=", "ElecPower"))))
oModule.ChangeProperty Array("NAME:AllTabs", Array("NAME:Trace", Array(_
  "NAME:PropServers", "Powers:Moving1.Speed*Moving1.Torque"), Array(_
  "NAME:ChangedProps", Array("NAME:Specify Name", "Value:=", true))))
oModule.ChangeProperty Array("NAME:AllTabs", Array("NAME:Trace", Array(_
  "NAME:PropServers", "Powers:Moving1.Speed*Moving1.Torque"), Array(_
  "NAME:ChangedProps", Array("NAME:Name", "Value:=", "MechPower"))))
oModule.AddTraceCharacteristics "Torque", "avg", Array(), Array("Specified", _
  "0.02s", "0.04s")
oModule.AddTraceCharacteristics "Currents", "rms", Array(), Array("Specified", _
  "0.02s", "0.04s")
oModule.AddTraceCharacteristics "Currents", "avg", Array(), Array("Specified", _
  "0.02s", "0.04s")
oModule.AddTraceCharacteristics "Voltages", "rms", Array(), Array("Specified", _
  "0.02s", "0.04s")
oModule.AddTraceCharacteristics "Induced Voltages", "rms", Array(), Array(_
  "Specified", "0.02s", "0.04s")
oModule.AddTraceCharacteristics "Powers", "avg", Array(), Array("Specified", _
  "0.02s", "0.04s")
oEditor.ShowWindow
Set oModule = oDesign.GetModule("OutputVariable")
oModule.CreateOutputVariable "pos", _
  "(Moving1.Position -13.425 * PI/180) * 16 + PI", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "cos0", "cos(pos)", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "cos1", "cos(pos-2*PI/3)", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "cos2", "cos(pos-4*PI/3)", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "sin0", "sin(pos)", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "sin1", "sin(pos-2*PI/3)", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "sin2", "sin(pos-4*PI/3)", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "Lad", _
  "L(PhaseA,PhaseA)*cos0 + L(PhaseA,PhaseB)*cos1 + L(PhaseA,PhaseC)*cos2", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Laq", _
  "L(PhaseA,PhaseA)*sin0 + L(PhaseA,PhaseB)*sin1 + L(PhaseA,PhaseC)*sin2", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Lbd", _
  "L(PhaseB,PhaseA)*cos0 + L(PhaseB,PhaseB)*cos1 + L(PhaseB,PhaseC)*cos2", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Lbq", _
  "L(PhaseB,PhaseA)*sin0 + L(PhaseB,PhaseB)*sin1 + L(PhaseB,PhaseC)*sin2", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Lcd", _
  "L(PhaseC,PhaseA)*cos0 + L(PhaseC,PhaseB)*cos1 + L(PhaseC,PhaseC)*cos2", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Lcq", _
  "L(PhaseC,PhaseA)*sin0 + L(PhaseC,PhaseB)*sin1 + L(PhaseC,PhaseC)*sin2", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "L_d", "(Lad*cos0 + Lbd*cos1 + Lcd*cos2) * 2/3", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "L_q", "(Laq*sin0 + Lbq*sin1 + Lcq*sin2) * 2/3", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Flux_d", _
  "(FluxLinkage(PhaseA)*cos0+FluxLinkage(PhaseB)*cos1+FluxLinkage(PhaseC)*cos2)*2/3", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Flux_q", _
  "(FluxLinkage(PhaseA)*sin0+FluxLinkage(PhaseB)*sin1+FluxLinkage(PhaseC)*sin2)*2/3", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "I_d", _
  "(Current(PhaseA)*cos0 + Current(PhaseB)*cos1 + Current(PhaseC)*cos2)*2/3", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "I_q", _
  "(Current(PhaseA)*sin0 + Current(PhaseB)*sin1 + Current(PhaseC)*sin2)*2/3", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Irms", "sqrt(I_d^2+I_q^2)/sqrt(2)", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Pcu", "3*Irms^2*0", "Setup1 : Transient", _
  "Transient", Array() 
Set oModule = oDesign.GetModule("ReportSetup")
oModule.CreateReport "L_dq", "Transient", "XY Plot", "Setup1 : Transient", _
  Array(), Array("Time:=", Array("All")), Array("X Component:=", "Time", _
  "Y Component:=", Array("L_d", "L_q")), Array()
oModule.CreateReport "Flux_dq", "Transient", "XY Plot", "Setup1 : Transient", _
  Array(), Array("Time:=", Array("All")), Array("X Component:=", "Time", _
  "Y Component:=", Array("Flux_d", "Flux_q")), Array()
oModule.CreateReport "I_dq", "Transient", "XY Plot", "Setup1 : Transient", _
  Array(), Array("Time:=", Array("All")), Array("X Component:=", "Time", _
  "Y Component:=", Array("I_d", "I_q")), Array()
oDesign.SetDesignSettings Array("NAME:Design Settings Data", _
  "ComputeTransientInductance:=", true, "ComputeIncrementalMatrix:=", false)
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.SetupYConnection Array(Array("NAME:YConnection", "Windings:=", _
  "PhaseA,PhaseB,PhaseC"))
