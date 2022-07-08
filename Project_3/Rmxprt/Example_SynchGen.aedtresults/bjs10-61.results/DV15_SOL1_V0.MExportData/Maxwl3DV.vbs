' ----------------------------------------------------------------------
' Script Created by RMxprt to generate Maxwell 3D Version 2016.0 design 
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
oDesign.SetSolutionType "Transient"
Set oModule = oDesign.GetModule("BoundarySetup")
if (oArgs.Count = 1) then 
oModule.EditExternalCircuit oArgs(0), Array(), Array(), Array(), Array()
end if
oEditor.SetModelValidationSettings Array("NAME:Validation Options", _
  "EntityCheckLevel:=", "Strict", "IgnoreUnclassifiedObjects:=", true)
oDesign.SetDesignSettings Array("NAME:Design Settings Data", _
  "InsulatorThreshold:=", 2.5e+06)
On Error Resume Next
Set oModule = oDesign.GetModule("MeshSetup")
oModule.InitialMeshSettings Array("NAME:MeshSettings", "MeshMethod:=", _
  "AnsoftTAU")
On Error Goto 0
On Error Resume Next
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:fractions", "PropType:=", "VariableProp", "UserDef:=", true, _
  "Value:=", "6"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:halfAxial", "PropType:=", "VariableProp", "UserDef:=", true, _
  "Value:=", "1"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:endRegion", "PropType:=", "VariableProp", "UserDef:=", true, _
  "Value:=", "501.27732402191668mm"))))
On Error Goto 0
Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.ModifyLibraries designName, Array("NAME:PersonalLib"), _
  Array("NAME:UserLib"), Array("NAME:SystemLib", "Materials:=", Array(_
  "Materials", "RMxprt"))
if (oDefinitionManager.DoesMaterialExist("M22_24G_3DSF0.900")) then
else
oDefinitionManager.AddMaterial Array("NAME:M22_24G_3DSF0.900", _
  "CoordinateSystemType:=", "Cartesian", Array("NAME:AttachedData"), Array(_
  "NAME:ModifierData"), Array("NAME:permeability", "property_type:=", _
  "nonlinear", "HUnit:=", "A_per_meter", "BUnit:=", "tesla", Array(_
  "NAME:BHCoordinates", Array("NAME:Coordinate", "X:=", 0, "Y:=", 0), Array(_
  "NAME:Coordinate", "X:=", 160, "Y:=", 0.8), Array("NAME:Coordinate", "X:=", _
  240, "Y:=", 1), Array("NAME:Coordinate", "X:=", 320, "Y:=", 1.1), Array(_
  "NAME:Coordinate", "X:=", 420, "Y:=", 1.2), Array("NAME:Coordinate", "X:=", _
  680, "Y:=", 1.3), Array("NAME:Coordinate", "X:=", 810, "Y:=", 1.325), _
  Array("NAME:Coordinate", "X:=", 1100, "Y:=", 1.375), Array(_
  "NAME:Coordinate", "X:=", 1400, "Y:=", 1.4), Array("NAME:Coordinate", "X:=", _
  2600, "Y:=", 1.475), Array("NAME:Coordinate", "X:=", 4200, "Y:=", 1.54), _
  Array("NAME:Coordinate", "X:=", 5087.91, "Y:=", 1.575), Array(_
  "NAME:Coordinate", "X:=", 6579.41, "Y:=", 1.625), Array("NAME:Coordinate", _
  "X:=", 8231.91, "Y:=", 1.675), Array("NAME:Coordinate", "X:=", 10258.9, _
  "Y:=", 1.725), Array("NAME:Coordinate", "X:=", 13348.9, "Y:=", 1.775), _
  Array("NAME:Coordinate", "X:=", 18788.9, "Y:=", 1.825), Array(_
  "NAME:Coordinate", "X:=", 28908.9, "Y:=", 1.875), Array("NAME:Coordinate", _
  "X:=", 46508.9, "Y:=", 1.925), Array("NAME:Coordinate", "X:=", 74308.9, _
  "Y:=", 1.975), Array("NAME:Coordinate", "X:=", 129350, "Y:=", 2.05), Array(_
  "NAME:Coordinate", "X:=", 248716, "Y:=", 2.2), Array("NAME:Coordinate", _
  "X:=", 447658, "Y:=", 2.45), Array("NAME:Coordinate", "X:=", 726178, "Y:=", _
  2.8), Array("NAME:Coordinate", "X:=", 3.51138e+06, "Y:=", 6.3))), Array(_
  "NAME:core_loss_type", "property_type:=", "ChoiceProperty", "Choice:=", _
  "Electrical Steel"), "core_loss_kh:=", 177.7, "core_loss_kc:=", 1.33, _
  "core_loss_ke:=", 1.22, "conductivity:=", 2e+06, "mass_density:=", 7650, _
  Array("NAME:stacking_type", "property_type:=", "ChoiceProperty", "Choice:=", _
  "Lamination"), "stacking_factor:=", "0.9", Array("NAME:stacking_direction", _
  "property_type:=", "ChoiceProperty", "Choice:=", "V(3)"))
end if
if (oDefinitionManager.DoesMaterialExist("steel_1008_3DSF0.960")) then
else
oDefinitionManager.AddMaterial Array("NAME:steel_1008_3DSF0.960", _
  "CoordinateSystemType:=", "Cartesian", Array("NAME:AttachedData"), Array(_
  "NAME:ModifierData"), Array("NAME:permeability", "property_type:=", _
  "nonlinear", "HUnit:=", "A_per_meter", "BUnit:=", "tesla", Array(_
  "NAME:BHCoordinates", Array("NAME:Coordinate", "X:=", 0, "Y:=", 0), Array(_
  "NAME:Coordinate", "X:=", 159.2, "Y:=", 0.2402), Array("NAME:Coordinate", _
  "X:=", 318.3, "Y:=", 0.8654), Array("NAME:Coordinate", "X:=", 477.5, "Y:=", _
  1.1106), Array("NAME:Coordinate", "X:=", 636.6, "Y:=", 1.2458), Array(_
  "NAME:Coordinate", "X:=", 795.8, "Y:=", 1.331), Array("NAME:Coordinate", _
  "X:=", 1591.5, "Y:=", 1.5), Array("NAME:Coordinate", "X:=", 3183.1, "Y:=", _
  1.6), Array("NAME:Coordinate", "X:=", 4774.6, "Y:=", 1.683), Array(_
  "NAME:Coordinate", "X:=", 6366.2, "Y:=", 1.741), Array("NAME:Coordinate", _
  "X:=", 7957.7, "Y:=", 1.78), Array("NAME:Coordinate", "X:=", 15915.5, "Y:=", _
  1.905), Array("NAME:Coordinate", "X:=", 31831, "Y:=", 2.025), Array(_
  "NAME:Coordinate", "X:=", 47746.5, "Y:=", 2.085), Array("NAME:Coordinate", _
  "X:=", 63662, "Y:=", 2.13), Array("NAME:Coordinate", "X:=", 79577.5, "Y:=", _
  2.165), Array("NAME:Coordinate", "X:=", 159155, "Y:=", 2.28), Array(_
  "NAME:Coordinate", "X:=", 318310, "Y:=", 2.485), Array("NAME:Coordinate", _
  "X:=", 397887, "Y:=", 2.5851), Array("NAME:Coordinate", "X:=", 1.19366e+06, _
  "Y:=", 3.5861))), "conductivity:=", 2e+06, "mass_density:=", 7872, Array(_
  "NAME:stacking_type", "property_type:=", "ChoiceProperty", "Choice:=", _
  "Lamination"), "stacking_factor:=", "0.96", Array("NAME:stacking_direction", _
  "property_type:=", "ChoiceProperty", "Choice:=", "V(3)"))
end if
if (oDefinitionManager.DoesMaterialExist("steel_1008_3DSF0.960")) then
else
oDefinitionManager.AddMaterial Array("NAME:steel_1008_3DSF0.960", _
  "CoordinateSystemType:=", "Cartesian", Array("NAME:AttachedData"), Array(_
  "NAME:ModifierData"), Array("NAME:permeability", "property_type:=", _
  "nonlinear", "HUnit:=", "A_per_meter", "BUnit:=", "tesla", Array(_
  "NAME:BHCoordinates", Array("NAME:Coordinate", "X:=", 0, "Y:=", 0), Array(_
  "NAME:Coordinate", "X:=", 159.2, "Y:=", 0.2402), Array("NAME:Coordinate", _
  "X:=", 318.3, "Y:=", 0.8654), Array("NAME:Coordinate", "X:=", 477.5, "Y:=", _
  1.1106), Array("NAME:Coordinate", "X:=", 636.6, "Y:=", 1.2458), Array(_
  "NAME:Coordinate", "X:=", 795.8, "Y:=", 1.331), Array("NAME:Coordinate", _
  "X:=", 1591.5, "Y:=", 1.5), Array("NAME:Coordinate", "X:=", 3183.1, "Y:=", _
  1.6), Array("NAME:Coordinate", "X:=", 4774.6, "Y:=", 1.683), Array(_
  "NAME:Coordinate", "X:=", 6366.2, "Y:=", 1.741), Array("NAME:Coordinate", _
  "X:=", 7957.7, "Y:=", 1.78), Array("NAME:Coordinate", "X:=", 15915.5, "Y:=", _
  1.905), Array("NAME:Coordinate", "X:=", 31831, "Y:=", 2.025), Array(_
  "NAME:Coordinate", "X:=", 47746.5, "Y:=", 2.085), Array("NAME:Coordinate", _
  "X:=", 63662, "Y:=", 2.13), Array("NAME:Coordinate", "X:=", 79577.5, "Y:=", _
  2.165), Array("NAME:Coordinate", "X:=", 159155, "Y:=", 2.28), Array(_
  "NAME:Coordinate", "X:=", 318310, "Y:=", 2.485), Array("NAME:Coordinate", _
  "X:=", 397887, "Y:=", 2.5851), Array("NAME:Coordinate", "X:=", 1.19366e+06, _
  "Y:=", 3.5861))), "conductivity:=", 2e+06, "mass_density:=", 7872, Array(_
  "NAME:stacking_type", "property_type:=", "ChoiceProperty", "Choice:=", _
  "Lamination"), "stacking_factor:=", "0.96", Array("NAME:stacking_direction", _
  "property_type:=", "ChoiceProperty", "Choice:=", "V(3)"))
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
  "Value:=", "1762.554648043833mm"), Array("NAME:Pair", "Name:=", "SegAngle", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Fractions", "Value:=", _
  "1"), Array("NAME:Pair", "Name:=", "HalfAxial", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "InfoCore", "Value:=", "0"))), Array(_
  "NAME:Attributes", "Name:=", "Band", "Flags:=", "", "Color:=", _
  "(0 255 255)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Band:CreateUserDefinedPart:1", _
  "Length", "840mm + 2*endRegion"
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/Band", "Version:=", "12.1", "NoOfParameters:=", 7, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "755.5mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "250mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "1762.554648043833mm"), Array("NAME:Pair", "Name:=", "SegAngle", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Fractions", "Value:=", _
  "1"), Array("NAME:Pair", "Name:=", "HalfAxial", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "InfoCore", "Value:=", "100"))), Array(_
  "NAME:Attributes", "Name:=", "Shaft", "Flags:=", "", "Color:=", _
  "(0 255 255)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Shaft:CreateUserDefinedPart:1", _
  "Length", "840mm + 2*endRegion"
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/Band", "Version:=", "12.1", "NoOfParameters:=", 7, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "755.5mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "1102.2mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "1762.554648043833mm"), Array("NAME:Pair", "Name:=", "SegAngle", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Fractions", "Value:=", _
  "6"), Array("NAME:Pair", "Name:=", "HalfAxial", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "InfoCore", "Value:=", "100"))), Array(_
  "NAME:Attributes", "Name:=", "OuterRegion", "Flags:=", "", "Color:=", _
  "(0 255 255)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion:CreateUserDefinedPart:1", "Length", "840mm + 2*endRegion"
On Error Goto 0
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion:CreateUserDefinedPart:1", "Fractions", "fractions"
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion:CreateUserDefinedPart:1", "HalfAxial", "halfAxial"
On Error Goto 0
oEditor.Copy Array("NAME:Selections", "Selections:=", "OuterRegion")
oEditor.Paste
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion1:CreateUserDefinedPart:1", "InfoCore", "2"
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "OuterRegion1"), _
  Array("NAME:ChangedProps", Array("NAME:Name", "Value:=", "MasterSheet"))))
oEditor.Copy Array("NAME:Selections", "Selections:=", "MasterSheet")
oEditor.Paste
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "MasterSheet1:CreateUserDefinedPart:1", "InfoCore", "3"
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "MasterSheet1"), _
  Array("NAME:ChangedProps", Array("NAME:Name", "Value:=", "SlaveSheet"))))
oEditor.Copy Array("NAME:Selections", "Selections:=", "OuterRegion")
oEditor.Paste
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion1:CreateUserDefinedPart:1", "InfoCore", "1"
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "OuterRegion1"), _
  Array("NAME:ChangedProps", Array("NAME:Name", "Value:=", "Tool"))))
oEditor.FitAll
Set oModule = oDesign.GetModule("ModelSetup")
oModule.SetSymmetryMultiplier "fractions*(1+halfAxial)"
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignMaster Array("NAME:Master", "Objects:=", Array("MasterSheet"), _
  Array("NAME:CoordSysVector", "Origin:=", Array("0mm", "0mm", "0mm"), _
  "UPos:=", Array("551.10000000000002mm", "0mm", "0mm")), "ReverseV:=", true)
oModule.AssignSlave Array("NAME:Slave", "Objects:=", Array("SlaveSheet"), _
  Array("NAME:CoordSysVector", "Origin:=", Array("0mm", "0mm", "0mm"), _
  "UPos:=", Array("275.55000000000007mm", "477.26660002560413mm", "0mm")), _
  "ReverseV:=", true, "Master:=", "Master", "RelationIsSame:=", false)
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/VentSlotCore", "Version:=", "12.1", "NoOfParameters:=", _
  27, "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "762mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "1102.2mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "760mm"), Array("NAME:Pair", "Name:=", "Skew", "Value:=", "0deg"), _
  Array("NAME:Pair", "Name:=", "Slots", "Value:=", "72"), Array("NAME:Pair", _
  "Name:=", "SlotType", "Value:=", "6"), Array("NAME:Pair", "Name:=", "Hs0", _
  "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs01", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "3.2mm"), Array("NAME:Pair", _
  "Name:=", "Hs2", "Value:=", "67mm"), Array("NAME:Pair", "Name:=", "Bs0", _
  "Value:=", "14.7mm"), Array("NAME:Pair", "Name:=", "Bs1", "Value:=", _
  "14.7mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", "14.7mm"), Array(_
  "NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", "HalfSlot", _
  "Value:=", "0"), Array("NAME:Pair", "Name:=", "VentHoles", "Value:=", "0"), _
  Array("NAME:Pair", "Name:=", "HoleDiaIn", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "HoleDiaOut", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "HoleLocIn", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "HoleLocOut", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "VentDucts", _
  "Value:=", "13"), Array("NAME:Pair", "Name:=", "DuctWidth", "Value:=", _
  "9.5mm"), Array("NAME:Pair", "Name:=", "DuctPitch", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", "30deg"), Array(_
  "NAME:Pair", "Name:=", "LenRegion", "Value:=", "1762.554648043833mm"), _
  Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", "0"))), Array(_
  "NAME:Attributes", "Name:=", "Stator", "Flags:=", "", "Color:=", _
  "(132 132 193)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "M22_24G_3DSF0.900", "SolveInside:=", true) 
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
  "Value:=", "760mm"), Array("NAME:Pair", "Name:=", "Skew", "Value:=", "0deg"), _
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
  "0.1mm"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", "10deg"), _
  Array("NAME:Pair", "Name:=", "LenRegion", "Value:=", "1762.554648043833mm"), _
  Array("NAME:Pair", "Name:=", "InfoCoil", "Value:=", "1"))), Array(_
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
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/LapCoil", "Version:=", "12.0", "NoOfParameters:=", 21, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "762mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "1102.2mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "760mm"), Array("NAME:Pair", "Name:=", "Skew", "Value:=", "0deg"), _
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
  "0.1mm"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", "10deg"), _
  Array("NAME:Pair", "Name:=", "LenRegion", "Value:=", "1762.554648043833mm"), _
  Array("NAME:Pair", "Name:=", "InfoCoil", "Value:=", "2"))), Array(_
  "NAME:Attributes", "Name:=", "CoilTerm", "Flags:=", "", "Color:=", _
  "(250 167 14)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "copper", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "CoilTerm:CreateUserDefinedPart:1", "LenRegion", "840mm + 2*endRegion"
On Error Goto 0
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", _
  "CoilTerm"), Array("NAME:DuplicateAroundAxisParameters", _
  "CoordinateSystemID:=", -1, "CreateNewObjects:=", true, "WhichAxis:=", "Z", _
  "AngleStr:=", "5deg", "NumClones:=", "72"), Array("NAME:Options", _
  "DuplicateBoundaries:=", false)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "CoilTerm"), _
  Array("NAME:ChangedProps", Array("NAME:Name", "Value:=", "CoilTerm_0"))))
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignWindingGroup Array("NAME:PhaseA", "Type:=", "External", _
  "IsSolid:=", false, "Current:=", "0A", "Resistance:=", "R1", "Inductance:=", _
  "Le1", "ParallelBranchesNum:=", "1")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignWindingGroup Array("NAME:PhaseB", "Type:=", "External", _
  "IsSolid:=", false, "Current:=", "0A", "Resistance:=", "R1", "Inductance:=", _
  "Le1", "ParallelBranchesNum:=", "1")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignWindingGroup Array("NAME:PhaseC", "Type:=", "External", _
  "IsSolid:=", false, "Current:=", "0A", "Resistance:=", "R1", "Inductance:=", _
  "Le1", "ParallelBranchesNum:=", "1")
oModule.AssignCoilTerminal Array("NAME:PhA_0", "Objects:=", Array("CoilTerm_0"), _
  "Conductor number:=", "conds", "Point out of terminal:=", false, _
  "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_1", "Objects:=", Array("CoilTerm_1"), _
  "Conductor number:=", "conds", "Point out of terminal:=", false, _
  "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_2", "Objects:=", Array("CoilTerm_2"), _
  "Conductor number:=", "conds", "Point out of terminal:=", false, _
  "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_3", "Objects:=", Array("CoilTerm_3"), _
  "Conductor number:=", "conds", "Point out of terminal:=", false, _
  "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhCRe_4", "Objects:=", Array(_
  "CoilTerm_4"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_5", "Objects:=", Array(_
  "CoilTerm_5"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_6", "Objects:=", Array(_
  "CoilTerm_6"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_7", "Objects:=", Array(_
  "CoilTerm_7"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhB_8", "Objects:=", Array("CoilTerm_8"), _
  "Conductor number:=", "conds", "Point out of terminal:=", false, _
  "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_9", "Objects:=", Array("CoilTerm_9"), _
  "Conductor number:=", "conds", "Point out of terminal:=", false, _
  "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_10", "Objects:=", Array(_
  "CoilTerm_10"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_11", "Objects:=", Array(_
  "CoilTerm_11"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhARe_12", "Objects:=", Array(_
  "CoilTerm_12"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_13", "Objects:=", Array(_
  "CoilTerm_13"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_14", "Objects:=", Array(_
  "CoilTerm_14"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_15", "Objects:=", Array(_
  "CoilTerm_15"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhC_16", "Objects:=", Array(_
  "CoilTerm_16"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_17", "Objects:=", Array(_
  "CoilTerm_17"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_18", "Objects:=", Array(_
  "CoilTerm_18"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_19", "Objects:=", Array(_
  "CoilTerm_19"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhBRe_20", "Objects:=", Array(_
  "CoilTerm_20"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_21", "Objects:=", Array(_
  "CoilTerm_21"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_22", "Objects:=", Array(_
  "CoilTerm_22"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_23", "Objects:=", Array(_
  "CoilTerm_23"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhA_24", "Objects:=", Array(_
  "CoilTerm_24"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_25", "Objects:=", Array(_
  "CoilTerm_25"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_26", "Objects:=", Array(_
  "CoilTerm_26"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_27", "Objects:=", Array(_
  "CoilTerm_27"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhCRe_28", "Objects:=", Array(_
  "CoilTerm_28"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_29", "Objects:=", Array(_
  "CoilTerm_29"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_30", "Objects:=", Array(_
  "CoilTerm_30"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_31", "Objects:=", Array(_
  "CoilTerm_31"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhB_32", "Objects:=", Array(_
  "CoilTerm_32"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_33", "Objects:=", Array(_
  "CoilTerm_33"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_34", "Objects:=", Array(_
  "CoilTerm_34"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_35", "Objects:=", Array(_
  "CoilTerm_35"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhARe_36", "Objects:=", Array(_
  "CoilTerm_36"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_37", "Objects:=", Array(_
  "CoilTerm_37"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_38", "Objects:=", Array(_
  "CoilTerm_38"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_39", "Objects:=", Array(_
  "CoilTerm_39"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhC_40", "Objects:=", Array(_
  "CoilTerm_40"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_41", "Objects:=", Array(_
  "CoilTerm_41"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_42", "Objects:=", Array(_
  "CoilTerm_42"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_43", "Objects:=", Array(_
  "CoilTerm_43"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhBRe_44", "Objects:=", Array(_
  "CoilTerm_44"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_45", "Objects:=", Array(_
  "CoilTerm_45"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_46", "Objects:=", Array(_
  "CoilTerm_46"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_47", "Objects:=", Array(_
  "CoilTerm_47"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhA_48", "Objects:=", Array(_
  "CoilTerm_48"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_49", "Objects:=", Array(_
  "CoilTerm_49"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_50", "Objects:=", Array(_
  "CoilTerm_50"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_51", "Objects:=", Array(_
  "CoilTerm_51"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhCRe_52", "Objects:=", Array(_
  "CoilTerm_52"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_53", "Objects:=", Array(_
  "CoilTerm_53"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_54", "Objects:=", Array(_
  "CoilTerm_54"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_55", "Objects:=", Array(_
  "CoilTerm_55"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhB_56", "Objects:=", Array(_
  "CoilTerm_56"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_57", "Objects:=", Array(_
  "CoilTerm_57"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_58", "Objects:=", Array(_
  "CoilTerm_58"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_59", "Objects:=", Array(_
  "CoilTerm_59"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhARe_60", "Objects:=", Array(_
  "CoilTerm_60"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_61", "Objects:=", Array(_
  "CoilTerm_61"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_62", "Objects:=", Array(_
  "CoilTerm_62"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_63", "Objects:=", Array(_
  "CoilTerm_63"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhC_64", "Objects:=", Array(_
  "CoilTerm_64"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_65", "Objects:=", Array(_
  "CoilTerm_65"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_66", "Objects:=", Array(_
  "CoilTerm_66"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_67", "Objects:=", Array(_
  "CoilTerm_67"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhBRe_68", "Objects:=", Array(_
  "CoilTerm_68"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_69", "Objects:=", Array(_
  "CoilTerm_69"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_70", "Objects:=", Array(_
  "CoilTerm_70"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_71", "Objects:=", Array(_
  "CoilTerm_71"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
Set oModule = oDesign.GetModule("MeshSetup")
oModule.AssignLengthOp Array("NAME:Length_Coil", "RefineInside:=", true, _
  "Objects:=", Array("Coil_0", "Coil_1", "Coil_2", "Coil_3", "Coil_4", _
  "Coil_5", "Coil_6", "Coil_7", "Coil_8", "Coil_9", "Coil_10", "Coil_11", _
  "Coil_12", "Coil_13", "Coil_14", "Coil_15", "Coil_16", "Coil_17", "Coil_18", _
  "Coil_19", "Coil_20", "Coil_21", "Coil_22", "Coil_23", "Coil_24", "Coil_25", _
  "Coil_26", "Coil_27", "Coil_28", "Coil_29", "Coil_30", "Coil_31", "Coil_32", _
  "Coil_33", "Coil_34", "Coil_35", "Coil_36", "Coil_37", "Coil_38", "Coil_39", _
  "Coil_40", "Coil_41", "Coil_42", "Coil_43", "Coil_44", "Coil_45", "Coil_46", _
  "Coil_47", "Coil_48", "Coil_49", "Coil_50", "Coil_51", "Coil_52", "Coil_53", _
  "Coil_54", "Coil_55", "Coil_56", "Coil_57", "Coil_58", "Coil_59", "Coil_60", _
  "Coil_61", "Coil_62", "Coil_63", "Coil_64", "Coil_65", "Coil_66", "Coil_67", _
  "Coil_68", "Coil_69", "Coil_70", "Coil_71"), "RestrictElem:=", false, _
  "NumMaxElem:=", "1000", "RestrictLength:=", true, "MaxLength:=", "29.4mm")
Set oModule = oDesign.GetModule("MeshSetup")
oModule.AssignTrueSurfOp Array("NAME:SurfApprox_Coil", "Objects:=", Array(_
  "Coil_0", "Coil_1", "Coil_2", "Coil_3", "Coil_4", "Coil_5", "Coil_6", _
  "Coil_7", "Coil_8", "Coil_9", "Coil_10", "Coil_11", "Coil_12", "Coil_13", _
  "Coil_14", "Coil_15", "Coil_16", "Coil_17", "Coil_18", "Coil_19", "Coil_20", _
  "Coil_21", "Coil_22", "Coil_23", "Coil_24", "Coil_25", "Coil_26", "Coil_27", _
  "Coil_28", "Coil_29", "Coil_30", "Coil_31", "Coil_32", "Coil_33", "Coil_34", _
  "Coil_35", "Coil_36", "Coil_37", "Coil_38", "Coil_39", "Coil_40", "Coil_41", _
  "Coil_42", "Coil_43", "Coil_44", "Coil_45", "Coil_46", "Coil_47", "Coil_48", _
  "Coil_49", "Coil_50", "Coil_51", "Coil_52", "Coil_53", "Coil_54", "Coil_55", _
  "Coil_56", "Coil_57", "Coil_58", "Coil_59", "Coil_60", "Coil_61", "Coil_62", _
  "Coil_63", "Coil_64", "Coil_65", "Coil_66", "Coil_67", "Coil_68", "Coil_69", _
  "Coil_70", "Coil_71"), "NormalDevChoice:=", 2, "NormalDev:=", "30deg", _
  "AspectRatioChoice:=", 1)
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "749mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "250mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "840mm"), Array("NAME:Pair", "Name:=", _
  "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", _
  "6"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array(_
  "NAME:Pair", "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "7mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "SlotPitch", "Value:=", "5.064deg"), Array(_
  "NAME:Pair", "Name:=", "CenterPitch", "Value:=", "14.736deg"), Array(_
  "NAME:Pair", "Name:=", "Poles", "Value:=", "6"), Array("NAME:Pair", _
  "Name:=", "WidthShoe", "Value:=", "280mm"), Array("NAME:Pair", "Name:=", _
  "HeightShoe", "Value:=", "53mm"), Array("NAME:Pair", "Name:=", "WidthBody", _
  "Value:=", "146mm"), Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", _
  "147mm"), Array("NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "Offset", "Value:=", "43mm"), Array("NAME:Pair", _
  "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", _
  "Value:=", "43mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", _
  "1"), Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "1762.554648043833mm"), Array("NAME:Pair", "Name:=", "InfoCore", _
  "Value:=", "0"))), Array("NAME:Attributes", "Name:=", "Rotor", "Flags:=", _
  "", "Color:=", "(132 132 193)", "Transparency:=", 0, _
  "PartCoordinateSystem:=", "Global", "MaterialName:=", _
  "steel_1008_3DSF0.960", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Rotor:CreateUserDefinedPart:1", _
  "LenRegion", "840mm + 2*endRegion"
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "749mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "250mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "840mm"), Array("NAME:Pair", "Name:=", _
  "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", _
  "6"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array(_
  "NAME:Pair", "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "7mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "SlotPitch", "Value:=", "5.064deg"), Array(_
  "NAME:Pair", "Name:=", "CenterPitch", "Value:=", "14.736deg"), Array(_
  "NAME:Pair", "Name:=", "Poles", "Value:=", "6"), Array("NAME:Pair", _
  "Name:=", "WidthShoe", "Value:=", "280mm"), Array("NAME:Pair", "Name:=", _
  "HeightShoe", "Value:=", "53mm"), Array("NAME:Pair", "Name:=", "WidthBody", _
  "Value:=", "146mm"), Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", _
  "147mm"), Array("NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "Offset", "Value:=", "43mm"), Array("NAME:Pair", _
  "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", _
  "Value:=", "43mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", _
  "1"), Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "1762.554648043833mm"), Array("NAME:Pair", "Name:=", "InfoCore", _
  "Value:=", "2"))), Array("NAME:Attributes", "Name:=", "Field", "Flags:=", _
  "", "Color:=", "(255 128 192)", "Transparency:=", 0, _
  "PartCoordinateSystem:=", "Global", "MaterialName:=", "copper", _
  "SolveInside:=", true) 
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
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "749mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "250mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "840mm"), Array("NAME:Pair", "Name:=", _
  "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", _
  "6"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array(_
  "NAME:Pair", "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "7mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "SlotPitch", "Value:=", "5.064deg"), Array(_
  "NAME:Pair", "Name:=", "CenterPitch", "Value:=", "14.736deg"), Array(_
  "NAME:Pair", "Name:=", "Poles", "Value:=", "6"), Array("NAME:Pair", _
  "Name:=", "WidthShoe", "Value:=", "280mm"), Array("NAME:Pair", "Name:=", _
  "HeightShoe", "Value:=", "53mm"), Array("NAME:Pair", "Name:=", "WidthBody", _
  "Value:=", "146mm"), Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", _
  "147mm"), Array("NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "Offset", "Value:=", "43mm"), Array("NAME:Pair", _
  "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", _
  "Value:=", "43mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", _
  "1"), Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "1762.554648043833mm"), Array("NAME:Pair", "Name:=", "InfoCore", _
  "Value:=", "4"))), Array("NAME:Attributes", "Name:=", "FieldTerm", _
  "Flags:=", "", "Color:=", "(255 128 192)", "Transparency:=", 0, _
  "PartCoordinateSystem:=", "Global", "MaterialName:=", "copper", _
  "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "FieldTerm:CreateUserDefinedPart:1", "LenRegion", "840mm + 2*endRegion"
On Error Goto 0
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", _
  "FieldTerm"), Array("NAME:DuplicateAroundAxisParameters", _
  "CoordinateSystemID:=", -1, "CreateNewObjects:=", true, "WhichAxis:=", "Z", _
  "AngleStr:=", "60deg", "NumClones:=", "6"), Array("NAME:Options", _
  "DuplicateBoundaries:=", false)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "FieldTerm"), _
  Array("NAME:ChangedProps", Array("NAME:Name", "Value:=", "FieldTerm_0"))))
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignWindingGroup Array("NAME:Field", "Type:=", "Current", _
  "IsSolid:=", false, "Current:=", "248.49A", "Resistance:=", "0ohm", _
  "Inductance:=", "0H", "ParallelBranchesNum:=", "1")
oModule.AssignCoilTerminal Array("NAME:FieldTerm_0", "Objects:=", Array(_
  "FieldTerm_0"), "Conductor number:=", 67.5, "Point out of terminal:=", _
  false, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTermRe_1", "Objects:=", Array(_
  "FieldTerm_1"), "Conductor number:=", 67.5, "Point out of terminal:=", _
  true, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTerm_2", "Objects:=", Array(_
  "FieldTerm_2"), "Conductor number:=", 67.5, "Point out of terminal:=", _
  false, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTermRe_3", "Objects:=", Array(_
  "FieldTerm_3"), "Conductor number:=", 67.5, "Point out of terminal:=", _
  true, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTerm_4", "Objects:=", Array(_
  "FieldTerm_4"), "Conductor number:=", 67.5, "Point out of terminal:=", _
  false, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTermRe_5", "Objects:=", Array(_
  "FieldTerm_5"), "Conductor number:=", 67.5, "Point out of terminal:=", _
  true, "Winding:=", "Field")
Set oModule = oDesign.GetModule("MeshSetup")
oModule.AssignLengthOp Array("NAME:Length_Field", "RefineInside:=", true, _
  "Objects:=", Array("Field_0", "Field_1", "Field_2", "Field_3", "Field_4", _
  "Field_5"), "RestrictElem:=", false, "NumMaxElem:=", "1000", _
  "RestrictLength:=", true, "MaxLength:=", "136.08mm")
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "749mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "250mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "840mm"), Array("NAME:Pair", "Name:=", _
  "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", _
  "6"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array(_
  "NAME:Pair", "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "7mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "SlotPitch", "Value:=", "5.064deg"), Array(_
  "NAME:Pair", "Name:=", "CenterPitch", "Value:=", "14.736deg"), Array(_
  "NAME:Pair", "Name:=", "Poles", "Value:=", "6"), Array("NAME:Pair", _
  "Name:=", "WidthShoe", "Value:=", "280mm"), Array("NAME:Pair", "Name:=", _
  "HeightShoe", "Value:=", "53mm"), Array("NAME:Pair", "Name:=", "WidthBody", _
  "Value:=", "146mm"), Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", _
  "147mm"), Array("NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "Offset", "Value:=", "43mm"), Array("NAME:Pair", _
  "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", _
  "Value:=", "43mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", _
  "1"), Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "1762.554648043833mm"), Array("NAME:Pair", "Name:=", "InfoCore", _
  "Value:=", "3"))), Array("NAME:Attributes", "Name:=", "Bar", "Flags:=", "", _
  "Color:=", "(119 198 206)", "Transparency:=", 0, "PartCoordinateSystem:=", _
  "Global", "MaterialName:=", "copper_115C", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Bar:CreateUserDefinedPart:1", _
  "LenRegion", "840mm + 2*endRegion"
On Error Goto 0
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Bar", _
  "Tool Parts:=", "Tool"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", true)
oEditor.SeparateBody  Array("NAME:Selections", "Selections:=", "Bar")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.SetEddyEffect Array("NAME:Eddy Effect Setting", Array(_
  "NAME:EddyEffectVector", Array("NAME:Data", "Object Name:=", "Bar", _
  "Eddy Effect:=", true)))
Set oModule = oDesign.GetModule("MeshSetup")
oModule.AssignLengthOp Array("NAME:Length_Bar", "RefineInside:=", true, _
  "Objects:=", Array("Bar"), "RestrictElem:=", false, "NumMaxElem:=", "1000", _
  "RestrictLength:=", true, "MaxLength:=", "28mm")
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "749mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "250mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "840mm"), Array("NAME:Pair", "Name:=", _
  "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", _
  "6"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array(_
  "NAME:Pair", "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "7mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "SlotPitch", "Value:=", "5.064deg"), Array(_
  "NAME:Pair", "Name:=", "CenterPitch", "Value:=", "14.736deg"), Array(_
  "NAME:Pair", "Name:=", "Poles", "Value:=", "6"), Array("NAME:Pair", _
  "Name:=", "WidthShoe", "Value:=", "280mm"), Array("NAME:Pair", "Name:=", _
  "HeightShoe", "Value:=", "53mm"), Array("NAME:Pair", "Name:=", "WidthBody", _
  "Value:=", "146mm"), Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", _
  "147mm"), Array("NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "Offset", "Value:=", "43mm"), Array("NAME:Pair", _
  "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", _
  "Value:=", "43mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", _
  "1"), Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "1762.554648043833mm"), Array("NAME:Pair", "Name:=", "InfoCore", _
  "Value:=", "100"))), Array("NAME:Attributes", "Name:=", "InnerRegion", _
  "Flags:=", "", "Color:=", "(0 255 255)", "Transparency:=", 0, _
  "PartCoordinateSystem:=", "Global", "MaterialName:=", "vacuum", _
  "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "InnerRegion:CreateUserDefinedPart:1", "LenRegion", "840mm + 2*endRegion"
On Error Goto 0
Set oModule = oDesign.GetModule("MeshSetup")
oModule.AssignLengthOp Array("NAME:Length_Main", "RefineInside:=", true, _
  "Objects:=", Array("Stator", "Rotor", "Band", "OuterRegion", "InnerRegion", _
  "Shaft"), "RestrictElem:=", false, "NumMaxElem:=", "1000", _
  "RestrictLength:=", true, "MaxLength:=", "141.28mm")
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "Band", _
  "OuterRegion", "InnerRegion", "MasterSheet", "SlaveSheet"), Array(_
  "NAME:ChangedProps", Array("NAME:Transparent", "Value:=", 0.75))))
oEditor.Move Array("NAME:Selections", "Selections:=", "CoilTerm_0,CoilTerm_1" & _
  ",CoilTerm_2,CoilTerm_3,CoilTerm_4,CoilTerm_5,CoilTerm_6,CoilTerm_7" & _
  ",CoilTerm_8,CoilTerm_9,CoilTerm_10,CoilTerm_11,CoilTerm_12,CoilTerm_13" & _
  ",CoilTerm_14,CoilTerm_15,CoilTerm_16,CoilTerm_17,CoilTerm_18" & _
  ",CoilTerm_19,CoilTerm_20,CoilTerm_21,CoilTerm_22,CoilTerm_23" & _
  ",CoilTerm_24,CoilTerm_25,CoilTerm_26,CoilTerm_27,CoilTerm_28" & _
  ",CoilTerm_29,CoilTerm_30,CoilTerm_31,CoilTerm_32,CoilTerm_33" & _
  ",CoilTerm_34,CoilTerm_35,CoilTerm_36,CoilTerm_37,CoilTerm_38" & _
  ",CoilTerm_39,CoilTerm_40,CoilTerm_41,CoilTerm_42,CoilTerm_43" & _
  ",CoilTerm_44,CoilTerm_45,CoilTerm_46,CoilTerm_47,CoilTerm_48" & _
  ",CoilTerm_49,CoilTerm_50,CoilTerm_51,CoilTerm_52,CoilTerm_53" & _
  ",CoilTerm_54,CoilTerm_55,CoilTerm_56,CoilTerm_57,CoilTerm_58" & _
  ",CoilTerm_59,CoilTerm_60,CoilTerm_61,CoilTerm_62,CoilTerm_63" & _
  ",CoilTerm_64,CoilTerm_65,CoilTerm_66,CoilTerm_67,CoilTerm_68" & _
  ",CoilTerm_69,CoilTerm_70,CoilTerm_71,FieldTerm_0,FieldTerm_1" & _
  ",FieldTerm_2,FieldTerm_3,FieldTerm_4,FieldTerm_5"), Array(_
  "NAME:TranslateParameters", "CoordinateSystemID:=", -1, _
  "TranslateVectorX:=", "0mm", "TranslateVectorY:=", "0mm", _
  "TranslateVectorZ:=", "190mm")
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Band,InnerRegion" & _
  ",Shaft,Stator,Coil_0,Coil_1,Coil_2,Coil_3,Coil_4,Coil_5,Coil_6,Coil_7" & _
  ",Coil_8,Coil_9,Coil_10,Coil_11,Coil_12,Coil_13,Coil_14,Coil_15,Coil_16" & _
  ",Coil_17,Coil_18,Coil_19,Coil_20,Coil_21,Coil_22,Coil_23,Coil_24" & _
  ",Coil_25,Coil_26,Coil_27,Coil_28,Coil_29,Coil_30,Coil_31,Coil_32" & _
  ",Coil_33,Coil_34,Coil_35,Coil_36,Coil_37,Coil_38,Coil_39,Coil_40" & _
  ",Coil_41,Coil_42,Coil_43,Coil_44,Coil_45,Coil_46,Coil_47,Coil_48" & _
  ",Coil_49,Coil_50,Coil_51,Coil_52,Coil_53,Coil_54,Coil_55,Coil_56" & _
  ",Coil_57,Coil_58,Coil_59,Coil_60,Coil_61,Coil_62,Coil_63,Coil_64" & _
  ",Coil_65,Coil_66,Coil_67,Coil_68,Coil_69,Coil_70,Coil_71,CoilTerm_0" & _
  ",CoilTerm_1,CoilTerm_2,CoilTerm_3,CoilTerm_4,CoilTerm_5,CoilTerm_6" & _
  ",CoilTerm_7,CoilTerm_8,CoilTerm_9,CoilTerm_10,CoilTerm_11,CoilTerm_12" & _
  ",CoilTerm_13,CoilTerm_14,CoilTerm_15,CoilTerm_16,CoilTerm_17" & _
  ",CoilTerm_18,CoilTerm_19,CoilTerm_20,CoilTerm_21,CoilTerm_22" & _
  ",CoilTerm_23,CoilTerm_24,CoilTerm_25,CoilTerm_26,CoilTerm_27" & _
  ",CoilTerm_28,CoilTerm_29,CoilTerm_30,CoilTerm_31,CoilTerm_32" & _
  ",CoilTerm_33,CoilTerm_34,CoilTerm_35,CoilTerm_36,CoilTerm_37" & _
  ",CoilTerm_38,CoilTerm_39,CoilTerm_40,CoilTerm_41,CoilTerm_42" & _
  ",CoilTerm_43,CoilTerm_44,CoilTerm_45,CoilTerm_46,CoilTerm_47" & _
  ",CoilTerm_48,CoilTerm_49,CoilTerm_50,CoilTerm_51,CoilTerm_52" & _
  ",CoilTerm_53,CoilTerm_54,CoilTerm_55,CoilTerm_56,CoilTerm_57" & _
  ",CoilTerm_58,CoilTerm_59,CoilTerm_60,CoilTerm_61,CoilTerm_62" & _
  ",CoilTerm_63,CoilTerm_64,CoilTerm_65,CoilTerm_66,CoilTerm_67" & _
  ",CoilTerm_68,CoilTerm_69,CoilTerm_70,CoilTerm_71,Rotor,Field_0,Field_1" & _
  ",Field_2,Field_3,Field_4,Field_5,FieldTerm_0,FieldTerm_1,FieldTerm_2" & _
  ",FieldTerm_3,FieldTerm_4,FieldTerm_5,Bar", "Tool Parts:=", "Tool"), _
  Array("NAME:SubtractParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
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
oModule.EditMotionSetup Array("NAME:Data", "Consider Mechanical Transient:=", _
  false)
Set oModule = oDesign.GetModule("AnalysisSetup")
oModule.InsertSetup "Transient", Array("NAME:Setup1", "StopTime:=", "0.2s", _
  "TimeStep:=", "0.0002s")
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
set oUnclassified = oEditor.GetObjectsInGroup("Unclassified")
Dim oObject
For Each oObject in oUnclassified
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", oObject), Array(_
  "NAME:ChangedProps", Array("NAME:Model", "Value:=", false))))
Next
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
