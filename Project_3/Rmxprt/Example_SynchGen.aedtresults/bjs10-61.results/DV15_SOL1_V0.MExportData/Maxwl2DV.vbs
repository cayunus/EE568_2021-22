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
  "Value:=", "3"))))
On Error Goto 0
Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.ModifyLibraries designName, Array("NAME:PersonalLib"), _
  Array("NAME:UserLib"), Array("NAME:SystemLib", "Materials:=", Array(_
  "Materials", "RMxprt"))
if (oDefinitionManager.DoesMaterialExist("M22_24G_2DSF0.781")) then
else
oDefinitionManager.AddMaterial Array("NAME:M22_24G_2DSF0.781", _
  "CoordinateSystemType:=", "Cartesian", Array("NAME:AttachedData"), Array(_
  "NAME:ModifierData"), Array("NAME:permeability", "property_type:=", _
  "nonlinear", "HUnit:=", "A_per_meter", "BUnit:=", "tesla", Array(_
  "NAME:BHCoordinates", Array("NAME:Coordinate", "X:=", 0, "Y:=", 0), Array(_
  "NAME:Coordinate", "X:=", 160, "Y:=", 0.624913), Array("NAME:Coordinate", _
  "X:=", 240, "Y:=", 0.781153), Array("NAME:Coordinate", "X:=", 320, "Y:=", _
  0.859283), Array("NAME:Coordinate", "X:=", 420, "Y:=", 0.93742), Array(_
  "NAME:Coordinate", "X:=", 680, "Y:=", 1.0156), Array("NAME:Coordinate", _
  "X:=", 810, "Y:=", 1.03516), Array("NAME:Coordinate", "X:=", 1100, "Y:=", _
  1.0743), Array("NAME:Coordinate", "X:=", 1400, "Y:=", 1.09391), Array(_
  "NAME:Coordinate", "X:=", 2600, "Y:=", 1.15282), Array("NAME:Coordinate", _
  "X:=", 4200, "Y:=", 1.20403), Array("NAME:Coordinate", "X:=", 5087.91, _
  "Y:=", 1.23161), Array("NAME:Coordinate", "X:=", 6579.41, "Y:=", 1.27108), _
  Array("NAME:Coordinate", "X:=", 8231.91, "Y:=", 1.31058), Array(_
  "NAME:Coordinate", "X:=", 10258.9, "Y:=", 1.3502), Array("NAME:Coordinate", _
  "X:=", 13348.9, "Y:=", 1.3901), Array("NAME:Coordinate", "X:=", 18788.9, _
  "Y:=", 1.43065), Array("NAME:Coordinate", "X:=", 28908.9, "Y:=", 1.47249), _
  Array("NAME:Coordinate", "X:=", 46508.9, "Y:=", 1.51639), Array(_
  "NAME:Coordinate", "X:=", 74308.9, "Y:=", 1.56309), Array("NAME:Coordinate", _
  "X:=", 129350, "Y:=", 1.63681), Array("NAME:Coordinate", "X:=", 248716, _
  "Y:=", 1.78681), Array("NAME:Coordinate", "X:=", 447658, "Y:=", 2.03681), _
  Array("NAME:Coordinate", "X:=", 726178, "Y:=", 2.38681), Array(_
  "NAME:Coordinate", "X:=", 3.51138e+06, "Y:=", 5.88681))), Array(_
  "NAME:core_loss_type", "property_type:=", "ChoiceProperty", "Choice:=", _
  "Electrical Steel"), "core_loss_kh:=", 227.504, "core_loss_kc:=", 1.70276, _
  "core_loss_ke:=", 1.38042, "conductivity:=", 2e+06, "mass_density:=", _
  5975.31) 
end if
if (oDefinitionManager.DoesMaterialExist("steel_1008_2DSF0.995")) then
else
oDefinitionManager.AddMaterial Array("NAME:steel_1008_2DSF0.995", _
  "CoordinateSystemType:=", "Cartesian", Array("NAME:AttachedData"), Array(_
  "NAME:ModifierData"), Array("NAME:permeability", "property_type:=", _
  "nonlinear", "HUnit:=", "A_per_meter", "BUnit:=", "tesla", Array(_
  "NAME:BHCoordinates", Array("NAME:Coordinate", "X:=", 0, "Y:=", 0), Array(_
  "NAME:Coordinate", "X:=", 159.2, "Y:=", 0.238956), Array("NAME:Coordinate", _
  "X:=", 318.3, "Y:=", 0.860917), Array("NAME:Coordinate", "X:=", 477.5, _
  "Y:=", 1.10485), Array("NAME:Coordinate", "X:=", 636.6, "Y:=", 1.23935), _
  Array("NAME:Coordinate", "X:=", 795.8, "Y:=", 1.32411), Array(_
  "NAME:Coordinate", "X:=", 1591.5, "Y:=", 1.49224), Array("NAME:Coordinate", _
  "X:=", 3183.1, "Y:=", 1.59173), Array("NAME:Coordinate", "X:=", 4774.6, _
  "Y:=", 1.67431), Array("NAME:Coordinate", "X:=", 6366.2, "Y:=", 1.73202), _
  Array("NAME:Coordinate", "X:=", 7957.7, "Y:=", 1.77083), Array(_
  "NAME:Coordinate", "X:=", 15915.5, "Y:=", 1.89523), Array("NAME:Coordinate", _
  "X:=", 31831, "Y:=", 2.01471), Array("NAME:Coordinate", "X:=", 47746.5, _
  "Y:=", 2.0745), Array("NAME:Coordinate", "X:=", 63662, "Y:=", 2.11937), _
  Array("NAME:Coordinate", "X:=", 79577.5, "Y:=", 2.1543), Array(_
  "NAME:Coordinate", "X:=", 159155, "Y:=", 2.26922), Array("NAME:Coordinate", _
  "X:=", 318310, "Y:=", 2.47419), Array("NAME:Coordinate", "X:=", 397887, _
  "Y:=", 2.57429), Array("NAME:Coordinate", "X:=", 1.19366e+06, "Y:=", _
  3.57529))), "conductivity:=", 2e+06, "mass_density:=", 7831.2) 
end if
if (oDefinitionManager.DoesMaterialExist("steel_1008_2DSF0.995")) then
else
oDefinitionManager.AddMaterial Array("NAME:steel_1008_2DSF0.995", _
  "CoordinateSystemType:=", "Cartesian", Array("NAME:AttachedData"), Array(_
  "NAME:ModifierData"), Array("NAME:permeability", "property_type:=", _
  "nonlinear", "HUnit:=", "A_per_meter", "BUnit:=", "tesla", Array(_
  "NAME:BHCoordinates", Array("NAME:Coordinate", "X:=", 0, "Y:=", 0), Array(_
  "NAME:Coordinate", "X:=", 159.2, "Y:=", 0.238956), Array("NAME:Coordinate", _
  "X:=", 318.3, "Y:=", 0.860917), Array("NAME:Coordinate", "X:=", 477.5, _
  "Y:=", 1.10485), Array("NAME:Coordinate", "X:=", 636.6, "Y:=", 1.23935), _
  Array("NAME:Coordinate", "X:=", 795.8, "Y:=", 1.32411), Array(_
  "NAME:Coordinate", "X:=", 1591.5, "Y:=", 1.49224), Array("NAME:Coordinate", _
  "X:=", 3183.1, "Y:=", 1.59173), Array("NAME:Coordinate", "X:=", 4774.6, _
  "Y:=", 1.67431), Array("NAME:Coordinate", "X:=", 6366.2, "Y:=", 1.73202), _
  Array("NAME:Coordinate", "X:=", 7957.7, "Y:=", 1.77083), Array(_
  "NAME:Coordinate", "X:=", 15915.5, "Y:=", 1.89523), Array("NAME:Coordinate", _
  "X:=", 31831, "Y:=", 2.01471), Array("NAME:Coordinate", "X:=", 47746.5, _
  "Y:=", 2.0745), Array("NAME:Coordinate", "X:=", 63662, "Y:=", 2.11937), _
  Array("NAME:Coordinate", "X:=", 79577.5, "Y:=", 2.1543), Array(_
  "NAME:Coordinate", "X:=", 159155, "Y:=", 2.26922), Array("NAME:Coordinate", _
  "X:=", 318310, "Y:=", 2.47419), Array("NAME:Coordinate", "X:=", 397887, _
  "Y:=", 2.57429), Array("NAME:Coordinate", "X:=", 1.19366e+06, "Y:=", _
  3.57529))), "conductivity:=", 2e+06, "mass_density:=", 7831.2) 
end if
if (oDefinitionManager.DoesMaterialExist("copper_115C")) then
else
oDefinitionManager.AddMaterial Array("NAME:copper_115C", _
  "CoordinateSystemType:=", "Cartesian", Array("NAME:AttachedData"), Array(_
  "NAME:ModifierData"), "conductivity:=", "4.07429e+07")
end if
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/Band", "Version:=", "12.1", "NoOfParameters:=", 7, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "755.5mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "250mm"), Array("NAME:Pair", "Name:=", "Length", _
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
  "Name:=", "DiaGap", "Value:=", "755.5mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "250mm"), Array("NAME:Pair", "Name:=", "Length", _
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
  "Name:=", "DiaGap", "Value:=", "755.5mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "1102.2mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", _
  "0deg"), Array("NAME:Pair", "Name:=", "Fractions", "Value:=", "3"), Array(_
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
  "OuterRegion", "XPosition:=", "275.55000000000007mm", "YPosition:=", _
  "477.26660002560413mm", "ZPosition:=", "0mm"))
oModule.AssignVectorPotential Array("NAME:VectorPotential1", "Edges:=", Array(edgeID), _
  "Value:=", "0", "CoordinateSystem:=", "")
edgeID = oEditor.GetEdgeByPosition(Array("NAME:Parameters", "BodyName:=", _
  "OuterRegion", "XPosition:=", "275.55000000000001mm", "YPosition:=", "0mm", _
  "ZPosition:=", "0mm"))
oModule.AssignMaster Array("NAME:Master1", "Edges:=", Array(edgeID), _
  "ReverseV:=", false)
edgeID = oEditor.GetEdgeByPosition(Array("NAME:Parameters", "BodyName:=", _
  "OuterRegion", "XPosition:=", "-137.77499999999995mm", "YPosition:=", _
  "238.6333000128021mm", "ZPosition:=", "0mm"))
oModule.AssignSlave Array("NAME:Slave1", "Edges:=", Array(edgeID), _
  "ReverseU:=", true, "Master:=", "Master1", "SameAsMaster:=", true)
oDesign.SetDesignSettings Array("NAME:Design Settings Data", "ModelDepth:=", _
  "733.401mm")
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SlotCore", "Version:=", "12.1", "NoOfParameters:=", 19, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "762mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "1102.2mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", "Value:=", "0deg"), _
  Array("NAME:Pair", "Name:=", "Slots", "Value:=", "72"), Array("NAME:Pair", _
  "Name:=", "SlotType", "Value:=", "6"), Array("NAME:Pair", "Name:=", "Hs0", _
  "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs01", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "3.2mm"), Array("NAME:Pair", _
  "Name:=", "Hs2", "Value:=", "67mm"), Array("NAME:Pair", "Name:=", "Bs0", _
  "Value:=", "14.7mm"), Array("NAME:Pair", "Name:=", "Bs1", "Value:=", _
  "14.7mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", "14.7mm"), Array(_
  "NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", "HalfSlot", _
  "Value:=", "0"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", _
  "15deg"), Array("NAME:Pair", "Name:=", "LenRegion", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", "0"))), Array(_
  "NAME:Attributes", "Name:=", "Stator", "Flags:=", "", "Color:=", _
  "(132 132 193)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "M22_24G_2DSF0.781", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Stator:CreateUserDefinedPart:1", _
  "LenRegion", "840mm + 2*endRegion"
On Error Goto 0
On Error Resume Next
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:delta", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "-28.514673505590345deg"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:conds", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "5"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:R1", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "0.2027877483906807ohm"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:Le1", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "0.0023052266650856533H"))))
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/LapCoil", "Version:=", "12.0", "NoOfParameters:=", 21, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "762mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "1102.2mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", "Value:=", "0deg"), _
  Array("NAME:Pair", "Name:=", "Slots", "Value:=", "72"), Array("NAME:Pair", _
  "Name:=", "SlotType", "Value:=", "6"), Array("NAME:Pair", "Name:=", "Hs0", _
  "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "3.2mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "67mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "14.7mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "14.7mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", _
  "14.7mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", _
  "Name:=", "Layers", "Value:=", "2"), Array("NAME:Pair", "Name:=", _
  "CoilPitch", "Value:=", "10"), Array("NAME:Pair", "Name:=", "EndExt", _
  "Value:=", "105mm"), Array("NAME:Pair", "Name:=", "SpanExt", "Value:=", _
  "0mm"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", "10deg"), _
  Array("NAME:Pair", "Name:=", "LenRegion", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "InfoCoil", "Value:=", "2"))), Array(_
  "NAME:Attributes", "Name:=", "Coil", "Flags:=", "", "Color:=", _
  "(250 167 14)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "copper", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Coil:CreateUserDefinedPart:1", _
  "LenRegion", "840mm + 2*endRegion"
On Error Goto 0
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "Coil"), _
  Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, _
  "CreateNewObjects:=", true, "WhichAxis:=", "Z", "AngleStr:=", "5deg", _
  "NumClones:=", "72"), Array("NAME:Options", "DuplicateBoundaries:=", false)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "Coil"), Array(_
  "NAME:ChangedProps", Array("NAME:Name", "Value:=", "Coil_0"))))
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_0,Coil_24,Coil_48"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_1,Coil_25,Coil_49"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_2,Coil_26,Coil_50"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_3,Coil_27,Coil_51"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_4,Coil_28,Coil_52"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_5,Coil_29,Coil_53"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_6,Coil_30,Coil_54"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_7,Coil_31,Coil_55"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_8,Coil_32,Coil_56"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_9,Coil_33,Coil_57"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_10,Coil_34,Coil_58"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_11,Coil_35,Coil_59"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_12,Coil_36,Coil_60"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_13,Coil_37,Coil_61"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_14,Coil_38,Coil_62"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_15,Coil_39,Coil_63"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_16,Coil_40,Coil_64"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_17,Coil_41,Coil_65"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_18,Coil_42,Coil_66"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_19,Coil_43,Coil_67"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_20,Coil_44,Coil_68"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_21,Coil_45,Coil_69"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_22,Coil_46,Coil_70"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_23,Coil_47,Coil_71"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/LapCoil", "Version:=", "12.0", "NoOfParameters:=", 21, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "762mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "1102.2mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", "Value:=", "0deg"), _
  Array("NAME:Pair", "Name:=", "Slots", "Value:=", "72"), Array("NAME:Pair", _
  "Name:=", "SlotType", "Value:=", "6"), Array("NAME:Pair", "Name:=", "Hs0", _
  "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "3.2mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "67mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "14.7mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "14.7mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", _
  "14.7mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", _
  "Name:=", "Layers", "Value:=", "2"), Array("NAME:Pair", "Name:=", _
  "CoilPitch", "Value:=", "10"), Array("NAME:Pair", "Name:=", "EndExt", _
  "Value:=", "105mm"), Array("NAME:Pair", "Name:=", "SpanExt", "Value:=", _
  "0mm"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", "10deg"), _
  Array("NAME:Pair", "Name:=", "LenRegion", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "InfoCoil", "Value:=", "3"))), Array(_
  "NAME:Attributes", "Name:=", "CoilRe", "Flags:=", "", "Color:=", _
  "(250 167 14)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "copper", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "CoilRe:CreateUserDefinedPart:1", _
  "LenRegion", "840mm + 2*endRegion"
On Error Goto 0
oEditor.Rotate Array("NAME:Selections", "Selections:=", "CoilRe"), Array(_
  "NAME:RotateParameters", "CoordinateSystemID:=", -1, "RotateAxis:=", "Z", _
  "RotateAngle:=", "-50deg")
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "CoilRe"), _
  Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, _
  "CreateNewObjects:=", true, "WhichAxis:=", "Z", "AngleStr:=", "5deg", _
  "NumClones:=", "72"), Array("NAME:Options", "DuplicateBoundaries:=", false)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "CoilRe"), Array(_
  "NAME:ChangedProps", Array("NAME:Name", "Value:=", "CoilRe_0"))))
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_0,CoilRe_24" & _
  ",CoilRe_48"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_1,CoilRe_25" & _
  ",CoilRe_49"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_2,CoilRe_26" & _
  ",CoilRe_50"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_3,CoilRe_27" & _
  ",CoilRe_51"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_4,CoilRe_28" & _
  ",CoilRe_52"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_5,CoilRe_29" & _
  ",CoilRe_53"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_6,CoilRe_30" & _
  ",CoilRe_54"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_7,CoilRe_31" & _
  ",CoilRe_55"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_8,CoilRe_32" & _
  ",CoilRe_56"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_9,CoilRe_33" & _
  ",CoilRe_57"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_10,CoilRe_34" & _
  ",CoilRe_58"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_11,CoilRe_35" & _
  ",CoilRe_59"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_12,CoilRe_36" & _
  ",CoilRe_60"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_13,CoilRe_37" & _
  ",CoilRe_61"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_14,CoilRe_38" & _
  ",CoilRe_62"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_15,CoilRe_39" & _
  ",CoilRe_63"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_16,CoilRe_40" & _
  ",CoilRe_64"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_17,CoilRe_41" & _
  ",CoilRe_65"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_18,CoilRe_42" & _
  ",CoilRe_66"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_19,CoilRe_43" & _
  ",CoilRe_67"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_20,CoilRe_44" & _
  ",CoilRe_68"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_21,CoilRe_45" & _
  ",CoilRe_69"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_22,CoilRe_46" & _
  ",CoilRe_70"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_23,CoilRe_47" & _
  ",CoilRe_71"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignWindingGroup Array("NAME:PhaseA", "Type:=", "Voltage", _
  "IsSolid:=", false, "Current:=", "0A", "Voltage:=", _
  "5388.88*sin(2*pi*50*time+delta)", "Resistance:=", "R1", "Inductance:=", _
  "Le1", "ParallelBranchesNum:=", "1")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignWindingGroup Array("NAME:PhaseB", "Type:=", "Voltage", _
  "IsSolid:=", false, "Current:=", "0A", "Voltage:=", _
  "5388.88*sin(2*pi*50*time+delta-2*pi/3)", "Resistance:=", "R1", _
  "Inductance:=", "Le1", "ParallelBranchesNum:=", "1")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignWindingGroup Array("NAME:PhaseC", "Type:=", "Voltage", _
  "IsSolid:=", false, "Current:=", "0A", "Voltage:=", _
  "5388.88*sin(2*pi*50*time+delta-4*pi/3)", "Resistance:=", "R1", _
  "Inductance:=", "Le1", "ParallelBranchesNum:=", "1")
oModule.AssignCoil Array("NAME:PhA_0", "Objects:=", Array("Coil_0"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_0", "Objects:=", Array("CoilRe_10"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_1", "Objects:=", Array("Coil_1"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_1", "Objects:=", Array("CoilRe_11"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_2", "Objects:=", Array("Coil_2"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_2", "Objects:=", Array("CoilRe_12"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_3", "Objects:=", Array("Coil_3"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_3", "Objects:=", Array("CoilRe_13"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhCRe_4", "Objects:=", Array("Coil_4"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_4", "Objects:=", Array("CoilRe_14"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_5", "Objects:=", Array("Coil_5"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_5", "Objects:=", Array("CoilRe_15"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_6", "Objects:=", Array("Coil_6"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_6", "Objects:=", Array("CoilRe_16"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_7", "Objects:=", Array("Coil_7"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_7", "Objects:=", Array("CoilRe_17"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhB_8", "Objects:=", Array("Coil_8"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_8", "Objects:=", Array("CoilRe_18"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_9", "Objects:=", Array("Coil_9"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_9", "Objects:=", Array("CoilRe_19"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_10", "Objects:=", Array("Coil_10"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_10", "Objects:=", Array("CoilRe_20"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_11", "Objects:=", Array("Coil_11"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_11", "Objects:=", Array("CoilRe_21"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhARe_12", "Objects:=", Array("Coil_12"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_12", "Objects:=", Array("CoilRe_22"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_13", "Objects:=", Array("Coil_13"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_13", "Objects:=", Array("CoilRe_23"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_14", "Objects:=", Array("Coil_14"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_15", "Objects:=", Array("Coil_15"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhC_16", "Objects:=", Array("Coil_16"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_17", "Objects:=", Array("Coil_17"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_18", "Objects:=", Array("Coil_18"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_19", "Objects:=", Array("Coil_19"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhBRe_20", "Objects:=", Array("Coil_20"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_21", "Objects:=", Array("Coil_21"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_22", "Objects:=", Array("Coil_22"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_23", "Objects:=", Array("Coil_23"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhA_62", "Objects:=", Array("CoilRe_0"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_63", "Objects:=", Array("CoilRe_1"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhCRe_64", "Objects:=", Array("CoilRe_2"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_65", "Objects:=", Array("CoilRe_3"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_66", "Objects:=", Array("CoilRe_4"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_67", "Objects:=", Array("CoilRe_5"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhB_68", "Objects:=", Array("CoilRe_6"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_69", "Objects:=", Array("CoilRe_7"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_70", "Objects:=", Array("CoilRe_8"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_71", "Objects:=", Array("CoilRe_9"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "749mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "250mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "6"), _
  Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs01", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", "14mm"), _
  Array("NAME:Pair", "Name:=", "Rs", "Value:=", "7mm"), Array("NAME:Pair", _
  "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", _
  "SlotPitch", "Value:=", "5.064deg"), Array("NAME:Pair", "Name:=", _
  "CenterPitch", "Value:=", "14.736deg"), Array("NAME:Pair", "Name:=", _
  "Poles", "Value:=", "6"), Array("NAME:Pair", "Name:=", "WidthShoe", _
  "Value:=", "280mm"), Array("NAME:Pair", "Name:=", "HeightShoe", "Value:=", _
  "53mm"), Array("NAME:Pair", "Name:=", "WidthBody", "Value:=", "146mm"), _
  Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", "147mm"), Array(_
  "NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Offset", "Value:=", "43mm"), Array("NAME:Pair", "Name:=", _
  "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Off2_y", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", "Value:=", _
  "43mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", "1"), _
  Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "15deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", "0"))), _
  Array("NAME:Attributes", "Name:=", "Rotor", "Flags:=", "", "Color:=", _
  "(132 132 193)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "steel_1008_2DSF0.995", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Rotor:CreateUserDefinedPart:1", _
  "LenRegion", "840mm + 2*endRegion"
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "749mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "250mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "6"), _
  Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs01", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", "14mm"), _
  Array("NAME:Pair", "Name:=", "Rs", "Value:=", "7mm"), Array("NAME:Pair", _
  "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", _
  "SlotPitch", "Value:=", "5.064deg"), Array("NAME:Pair", "Name:=", _
  "CenterPitch", "Value:=", "14.736deg"), Array("NAME:Pair", "Name:=", _
  "Poles", "Value:=", "6"), Array("NAME:Pair", "Name:=", "WidthShoe", _
  "Value:=", "280mm"), Array("NAME:Pair", "Name:=", "HeightShoe", "Value:=", _
  "53mm"), Array("NAME:Pair", "Name:=", "WidthBody", "Value:=", "146mm"), _
  Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", "147mm"), Array(_
  "NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Offset", "Value:=", "43mm"), Array("NAME:Pair", "Name:=", _
  "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Off2_y", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", "Value:=", _
  "43mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", "1"), _
  Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "15deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", "4"))), _
  Array("NAME:Attributes", "Name:=", "Field", "Flags:=", "", "Color:=", _
  "(255 128 192)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "copper", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Field:CreateUserDefinedPart:1", _
  "LenRegion", "840mm + 2*endRegion"
On Error Goto 0
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "Field"), _
  Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, _
  "CreateNewObjects:=", true, "WhichAxis:=", "Z", "AngleStr:=", "60deg", _
  "NumClones:=", "6"), Array("NAME:Options", "DuplicateBoundaries:=", false)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "Field"), Array(_
  "NAME:ChangedProps", Array("NAME:Name", "Value:=", "Field_0"))))
oEditor.Unite Array("NAME:Selections", "Selections:=", "Field_0,Field_2,Field_4"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Field_1,Field_3,Field_5"), _
  Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "749mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "250mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "6"), _
  Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs01", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", "14mm"), _
  Array("NAME:Pair", "Name:=", "Rs", "Value:=", "7mm"), Array("NAME:Pair", _
  "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", _
  "SlotPitch", "Value:=", "5.064deg"), Array("NAME:Pair", "Name:=", _
  "CenterPitch", "Value:=", "14.736deg"), Array("NAME:Pair", "Name:=", _
  "Poles", "Value:=", "6"), Array("NAME:Pair", "Name:=", "WidthShoe", _
  "Value:=", "280mm"), Array("NAME:Pair", "Name:=", "HeightShoe", "Value:=", _
  "53mm"), Array("NAME:Pair", "Name:=", "WidthBody", "Value:=", "146mm"), _
  Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", "147mm"), Array(_
  "NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Offset", "Value:=", "43mm"), Array("NAME:Pair", "Name:=", _
  "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Off2_y", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", "Value:=", _
  "43mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", "1"), _
  Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "15deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", "5"))), _
  Array("NAME:Attributes", "Name:=", "FieldRe", "Flags:=", "", "Color:=", _
  "(255 128 192)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "copper", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "FieldRe:CreateUserDefinedPart:1", _
  "LenRegion", "840mm + 2*endRegion"
On Error Goto 0
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "FieldRe"), _
  Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, _
  "CreateNewObjects:=", true, "WhichAxis:=", "Z", "AngleStr:=", "60deg", _
  "NumClones:=", "6"), Array("NAME:Options", "DuplicateBoundaries:=", false)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "FieldRe"), Array(_
  "NAME:ChangedProps", Array("NAME:Name", "Value:=", "FieldRe_0"))))
oEditor.Unite Array("NAME:Selections", "Selections:=", "FieldRe_0,FieldRe_2" & _
  ",FieldRe_4"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "FieldRe_1,FieldRe_3" & _
  ",FieldRe_5"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignWindingGroup Array("NAME:Field", "Type:=", "Current", _
  "IsSolid:=", false, "Current:=", "248.49A", "Resistance:=", "0ohm", _
  "Inductance:=", "0H", "ParallelBranchesNum:=", "1")
oModule.AssignCoil Array("NAME:Field_0", "Objects:=", Array("Field_0"), _
  "Conductor number:=", 67.5, "PolarityType:=", "Positive", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:FieldRe_0", "Objects:=", Array("FieldRe_0"), _
  "Conductor number:=", 67.5, "PolarityType:=", "Negative", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:FieldRe_1", "Objects:=", Array("Field_1"), _
  "Conductor number:=", 67.5, "PolarityType:=", "Negative", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:Field_1", "Objects:=", Array("FieldRe_1"), _
  "Conductor number:=", 67.5, "PolarityType:=", "Positive", "Winding:=", _
  "Field")
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "749mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "250mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "6"), _
  Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs01", _
  "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", "14mm"), _
  Array("NAME:Pair", "Name:=", "Rs", "Value:=", "7mm"), Array("NAME:Pair", _
  "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", _
  "SlotPitch", "Value:=", "5.064deg"), Array("NAME:Pair", "Name:=", _
  "CenterPitch", "Value:=", "14.736deg"), Array("NAME:Pair", "Name:=", _
  "Poles", "Value:=", "6"), Array("NAME:Pair", "Name:=", "WidthShoe", _
  "Value:=", "280mm"), Array("NAME:Pair", "Name:=", "HeightShoe", "Value:=", _
  "53mm"), Array("NAME:Pair", "Name:=", "WidthBody", "Value:=", "146mm"), _
  Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", "147mm"), Array(_
  "NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Offset", "Value:=", "43mm"), Array("NAME:Pair", "Name:=", _
  "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Off2_y", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", "Value:=", _
  "43mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", "1"), _
  Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "15deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", "3"))), _
  Array("NAME:Attributes", "Name:=", "Bar", "Flags:=", "", "Color:=", _
  "(119 198 206)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "copper_115C", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Bar:CreateUserDefinedPart:1", _
  "LenRegion", "840mm + 2*endRegion"
On Error Goto 0
oEditor.Split Array("NAME:Selections", "Selections:=", "Bar", _
  "NewPartsModelFlag:=", "Model"), Array("NAME:SplitToParameters", _
  "CoordinateSystemID:=", -1, "SplitPlane:=", "ZX", "WhichSide:=", _
  "PositiveOnly", "SplitCrossingObjectsOnly:=", true)
oEditor.Rotate Array("NAME:Selections", "Selections:=", "Bar"), Array(_
  "NAME:RotateParameters", "CoordinateSystemID:=", -1, "RotateAxis:=", "Z", _
  "RotateAngle:=", "120deg")
oEditor.Split Array("NAME:Selections", "Selections:=", "Bar", _
  "NewPartsModelFlag:=", "Model"), Array("NAME:SplitToParameters", _
  "CoordinateSystemID:=", -1, "SplitPlane:=", "ZX", "WhichSide:=", _
  "PositiveOnly", "SplitCrossingObjectsOnly:=", true)
oEditor.Rotate Array("NAME:Selections", "Selections:=", "Bar"), Array(_
  "NAME:RotateParameters", "CoordinateSystemID:=", -1, "RotateAxis:=", "Z", _
  "RotateAngle:=", "-120deg")
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "Bar"), _
  Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, _
  "CreateNewObjects:=", true, "WhichAxis:=", "Z", "AngleStr:=", "60deg", _
  "NumClones:=", "2"), Array("NAME:Options", "DuplicateBoundaries:=", false)
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Bar", _
  "Tool Parts:=", "Tool"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", true)
oEditor.SeparateBody  Array("NAME:Selections", "Selections:=", "Bar")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignEndConnection Array("NAME:EndConnection1", "Objects:=", Array(_
  "Bar", "Bar_Separate1", "Bar_Separate2", "Bar_Separate3", "Bar_Separate4", _
  "Bar_Separate5"), "ResistanceValue:=", "0ohm", "InductanceValue:=", "0H")
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Bar_1", _
  "Tool Parts:=", "Tool"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", true)
oEditor.SeparateBody  Array("NAME:Selections", "Selections:=", "Bar_1")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignEndConnection Array("NAME:EndConnection2", "Objects:=", Array(_
  "Bar_1", "Bar_1_Separate1", "Bar_1_Separate2", "Bar_1_Separate3", _
  "Bar_1_Separate4", "Bar_1_Separate5"), "ResistanceValue:=", "0ohm", _
  "InductanceValue:=", "0H")
Set oModule = oDesign.GetModule("MeshSetup")
oModule.AssignTrueSurfOp Array("NAME:SurfApprox_Bar", "Objects:=", Array("Bar", _
  "Bar_Separate1", "Bar_Separate2", "Bar_Separate3", "Bar_Separate4", _
  "Bar_Separate5", "Bar_1", "Bar_1_Separate1", "Bar_1_Separate2", _
  "Bar_1_Separate3", "Bar_1_Separate4", "Bar_1_Separate5"), "SurfDevChoice:=", _
  2, "SurfDev:=", "0.3745mm", "NormalDevChoice:=", 2, "NormalDev:=", "15deg", _
  "AspectRatioChoice:=", 1)
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "749mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "250mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "6"), _
  Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs01", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", "14mm"), _
  Array("NAME:Pair", "Name:=", "Rs", "Value:=", "7mm"), Array("NAME:Pair", _
  "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", _
  "SlotPitch", "Value:=", "5.064deg"), Array("NAME:Pair", "Name:=", _
  "CenterPitch", "Value:=", "14.736deg"), Array("NAME:Pair", "Name:=", _
  "Poles", "Value:=", "6"), Array("NAME:Pair", "Name:=", "WidthShoe", _
  "Value:=", "280mm"), Array("NAME:Pair", "Name:=", "HeightShoe", "Value:=", _
  "53mm"), Array("NAME:Pair", "Name:=", "WidthBody", "Value:=", "146mm"), _
  Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", "147mm"), Array(_
  "NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Offset", "Value:=", "43mm"), Array("NAME:Pair", "Name:=", _
  "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Off2_y", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", "Value:=", _
  "43mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", "1"), _
  Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "15deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", _
  "100"))), Array("NAME:Attributes", "Name:=", "InnerRegion", "Flags:=", "", _
  "Color:=", "(0 255 255)", "Transparency:=", 0, "PartCoordinateSystem:=", _
  "Global", "MaterialName:=", "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "InnerRegion:CreateUserDefinedPart:1", "LenRegion", "840mm + 2*endRegion"
On Error Goto 0
Set oModule = oDesign.GetModule("MeshSetup")
oModule.AssignTrueSurfOp Array("NAME:SurfApprox_Main", "Objects:=", Array(_
  "Stator", "Rotor", "Band", "OuterRegion", "InnerRegion", "Shaft"), _
  "SurfDevChoice:=", 2, "SurfDev:=", "0.5511mm", "NormalDevChoice:=", 2, _
  "NormalDev:=", "15deg", "AspectRatioChoice:=", 1)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "Band", _
  "OuterRegion", "InnerRegion"), Array("NAME:ChangedProps", Array(_
  "NAME:Transparent", "Value:=", 0.75))))
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Band,InnerRegion" & _
  ",Shaft,Stator,Coil_0,Coil_1,Coil_2,Coil_3,Coil_4,Coil_5,Coil_6,Coil_7" & _
  ",Coil_8,Coil_9,Coil_10,Coil_11,Coil_12,Coil_13,Coil_14,Coil_15,Coil_16" & _
  ",Coil_17,Coil_18,Coil_19,Coil_20,Coil_21,Coil_22,Coil_23,CoilRe_0" & _
  ",CoilRe_1,CoilRe_2,CoilRe_3,CoilRe_4,CoilRe_5,CoilRe_6,CoilRe_7" & _
  ",CoilRe_8,CoilRe_9,CoilRe_10,CoilRe_11,CoilRe_12,CoilRe_13,CoilRe_14" & _
  ",CoilRe_15,CoilRe_16,CoilRe_17,CoilRe_18,CoilRe_19,CoilRe_20,CoilRe_21" & _
  ",CoilRe_22,CoilRe_23,Rotor,Field_0,Field_1,FieldRe_0,FieldRe_1", _
  "Tool Parts:=", "Tool"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.FitAll
Set oModule = oDesign.GetModule("ModelSetup")
oModule.AssignBand Array("NAME:MotionSetup1", "Move Type:=", "Rotate", _
  "Coordinate System:=", "Global", "Axis:=", "Z", "Is Positive:=", true, _
  "InitPos:=", "65deg", "HasRotateLimit:=", false, "NonCylindrical:=", _
  false, "Consider Mechanical Transient:=", true, "Angular Velocity:=", _
  "1000rpm", "Moment of Inertia:=", 183.162, "Damping:=", 2.70901, _
  "Load Torque:=", _
  "if(speed<104.72, -175.319*speed, -1.92259e+06/speed) - 175.319*(speed-104.72)*10", _
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
oModule.CreateOutputVariable "pos", "(Moving1.Position -65 * PI/180) * 3 + PI", _
  "Setup1 : Transient", "Transient", Array() 
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
