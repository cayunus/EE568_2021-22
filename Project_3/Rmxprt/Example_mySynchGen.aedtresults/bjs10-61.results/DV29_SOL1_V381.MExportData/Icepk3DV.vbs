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
if (oArgs.Count > 0) then 
  Set oDesign = oProject.GetDesign(oArgs(0))
else
  Set oDesign = oProject.GetActiveDesign()
end if
designName = oDesign.GetName
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.SetModelUnits Array("NAME:Units Parameter", "Units:=", "mm", _
  "Rescale:=", false)
oDesign.SetSolutionType "Transient"
Set oModule = oDesign.GetModule("BoundarySetup")
oEditor.SetModelValidationSettings Array("NAME:Validation Options", _
  "EntityCheckLevel:=", "Strict", "IgnoreUnclassifiedObjects:=", true)
oDesign.SetDesignSettings Array("NAME:Design Settings Data", _
  "InsulatorThreshold:=", 2.5e+06)
On Error Resume Next
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:fractions", "PropType:=", "VariableProp", "UserDef:=", true, _
  "Value:=", "1"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:halfAxial", "PropType:=", "VariableProp", "UserDef:=", true, _
  "Value:=", "0"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:endRegion", "PropType:=", "VariableProp", "UserDef:=", true, _
  "Value:=", "678.85297560927211mm"))))
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
if (oDefinitionManager.DoesMaterialExist("M22_24G_3DSF0.960")) then
else
oDefinitionManager.AddMaterial Array("NAME:M22_24G_3DSF0.960", _
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
  "Lamination"), "stacking_factor:=", "0.96", Array("NAME:stacking_direction", _
  "property_type:=", "ChoiceProperty", "Choice:=", "V(3)"))
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
  "Value:=", "2423.705951218544mm"), Array("NAME:Pair", "Name:=", "SegAngle", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Fractions", "Value:=", _
  "1"), Array("NAME:Pair", "Name:=", "HalfAxial", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "InfoCore", "Value:=", "100"))), Array(_
  "NAME:Attributes", "Name:=", "Shaft", "Flags:=", "", "Color:=", _
  "(0 255 255)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Shaft:CreateUserDefinedPart:1", _
  "Length", "1180mm + 2*endRegion"
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/Band", "Version:=", "12.1", "NoOfParameters:=", 7, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5414.5mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "5974mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "2423.705951218544mm"), Array("NAME:Pair", "Name:=", "SegAngle", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Fractions", "Value:=", _
  "1"), Array("NAME:Pair", "Name:=", "HalfAxial", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "InfoCore", "Value:=", "100"))), Array(_
  "NAME:Attributes", "Name:=", "OuterRegion", "Flags:=", "", "Color:=", _
  "(0 255 255)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion:CreateUserDefinedPart:1", "Length", "1180mm + 2*endRegion"
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
  "OuterRegion1:CreateUserDefinedPart:1", "InfoCore", "1"
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "OuterRegion1"), _
  Array("NAME:ChangedProps", Array("NAME:Name", "Value:=", "Tool"))))
oEditor.FitAll
Set oModule = oDesign.GetModule("ModelSetup")
oModule.SetSymmetryMultiplier "fractions*(1+halfAxial)"
Set oModule = oDesign.GetModule("BoundarySetup")
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/VentSlotCore", "Version:=", "12.1", "NoOfParameters:=", _
  27, "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5431mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "5974mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "1066mm"), Array("NAME:Pair", "Name:=", "Skew", "Value:=", _
  "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "300"), Array(_
  "NAME:Pair", "Name:=", "SlotType", "Value:=", "6"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "1mm"), Array("NAME:Pair", "Name:=", "Hs01", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "4.2mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "120mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "22.7mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "22.7mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", _
  "22.7mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", _
  "Name:=", "HalfSlot", "Value:=", "0"), Array("NAME:Pair", "Name:=", _
  "VentHoles", "Value:=", "0"), Array("NAME:Pair", "Name:=", "HoleDiaIn", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "HoleDiaOut", "Value:=", _
  "0mm"), Array("NAME:Pair", "Name:=", "HoleLocIn", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "HoleLocOut", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "VentDucts", "Value:=", "10"), Array("NAME:Pair", "Name:=", _
  "DuctWidth", "Value:=", "9.5mm"), Array("NAME:Pair", "Name:=", "DuctPitch", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", _
  "30deg"), Array("NAME:Pair", "Name:=", "LenRegion", "Value:=", _
  "2423.705951218544mm"), Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", _
  "0"))), Array("NAME:Attributes", "Name:=", "Stator", "Flags:=", "", _
  "Color:=", "(132 132 193)", "Transparency:=", 0, "PartCoordinateSystem:=", _
  "Global", "MaterialName:=", "M22_24G_3DSF0.900", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Stator:CreateUserDefinedPart:1", _
  "LenRegion", "1180mm + 2*endRegion"
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/VentSlotCore", "Version:=", "12.1", "NoOfParameters:=", _
  27, "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5431mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "5974mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "1066mm"), Array("NAME:Pair", "Name:=", "Skew", "Value:=", _
  "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "300"), Array(_
  "NAME:Pair", "Name:=", "SlotType", "Value:=", "6"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "1mm"), Array("NAME:Pair", "Name:=", "Hs01", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "4.2mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "120mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "22.7mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "22.7mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", _
  "22.7mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", _
  "Name:=", "HalfSlot", "Value:=", "0"), Array("NAME:Pair", "Name:=", _
  "VentHoles", "Value:=", "0"), Array("NAME:Pair", "Name:=", "HoleDiaIn", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "HoleDiaOut", "Value:=", _
  "0mm"), Array("NAME:Pair", "Name:=", "HoleLocIn", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "HoleLocOut", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "VentDucts", "Value:=", "10"), Array("NAME:Pair", "Name:=", _
  "DuctWidth", "Value:=", "9.5mm"), Array("NAME:Pair", "Name:=", "DuctPitch", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", _
  "30deg"), Array("NAME:Pair", "Name:=", "LenRegion", "Value:=", _
  "2423.705951218544mm"), Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", _
  "1"))), Array("NAME:Attributes", "Name:=", "StatorSlotInsu", "Flags:=", "", _
  "Color:=", "(132 132 193)", "Transparency:=", 0, "PartCoordinateSystem:=", _
  "Global", "MaterialName:=", "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "StatorSlotInsu:CreateUserDefinedPart:1", "LenRegion", _
  "1180mm + 2*endRegion"
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
  "Value:=", "1066mm"), Array("NAME:Pair", "Name:=", "Skew", "Value:=", _
  "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "300"), Array(_
  "NAME:Pair", "Name:=", "SlotType", "Value:=", "6"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "1mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "4.2mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "120mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "22.7mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "22.7mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "22.7mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "Layers", "Value:=", "2"), Array("NAME:Pair", _
  "Name:=", "CoilPitch", "Value:=", "9"), Array("NAME:Pair", "Name:=", _
  "EndExt", "Value:=", "150mm"), Array("NAME:Pair", "Name:=", "SpanExt", _
  "Value:=", "0.1mm"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", _
  "10deg"), Array("NAME:Pair", "Name:=", "LenRegion", "Value:=", _
  "2423.705951218544mm"), Array("NAME:Pair", "Name:=", "InfoCoil", "Value:=", _
  "1"))), Array("NAME:Attributes", "Name:=", "Coil", "Flags:=", "", "Color:=", _
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
On Error Resume Next
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "StatorSlotInsu", _
  "Tool Parts:=", "Coil_0,Coil_1,Coil_2,Coil_3,Coil_4,Coil_5,Coil_6,Coil_7" & _
  ",Coil_8,Coil_9,Coil_10,Coil_11,Coil_12,Coil_13,Coil_14,Coil_15,Coil_16" & _
  ",Coil_17,Coil_18,Coil_19,Coil_20,Coil_21,Coil_22,Coil_23,Coil_24" & _
  ",Coil_25,Coil_26,Coil_27,Coil_28,Coil_29,Coil_30,Coil_31,Coil_32" & _
  ",Coil_33,Coil_34,Coil_35,Coil_36,Coil_37,Coil_38,Coil_39,Coil_40" & _
  ",Coil_41,Coil_42,Coil_43,Coil_44,Coil_45,Coil_46,Coil_47,Coil_48" & _
  ",Coil_49,Coil_50,Coil_51,Coil_52,Coil_53,Coil_54,Coil_55,Coil_56" & _
  ",Coil_57,Coil_58,Coil_59,Coil_60,Coil_61,Coil_62,Coil_63,Coil_64" & _
  ",Coil_65,Coil_66,Coil_67,Coil_68,Coil_69,Coil_70,Coil_71,Coil_72" & _
  ",Coil_73,Coil_74,Coil_75,Coil_76,Coil_77,Coil_78,Coil_79,Coil_80" & _
  ",Coil_81,Coil_82,Coil_83,Coil_84,Coil_85,Coil_86,Coil_87,Coil_88" & _
  ",Coil_89,Coil_90,Coil_91,Coil_92,Coil_93,Coil_94,Coil_95,Coil_96" & _
  ",Coil_97,Coil_98,Coil_99,Coil_100,Coil_101,Coil_102,Coil_103,Coil_104" & _
  ",Coil_105,Coil_106,Coil_107,Coil_108,Coil_109,Coil_110,Coil_111" & _
  ",Coil_112,Coil_113,Coil_114,Coil_115,Coil_116,Coil_117,Coil_118" & _
  ",Coil_119,Coil_120,Coil_121,Coil_122,Coil_123,Coil_124,Coil_125" & _
  ",Coil_126,Coil_127,Coil_128,Coil_129,Coil_130,Coil_131,Coil_132" & _
  ",Coil_133,Coil_134,Coil_135,Coil_136,Coil_137,Coil_138,Coil_139" & _
  ",Coil_140,Coil_141,Coil_142,Coil_143,Coil_144,Coil_145,Coil_146" & _
  ",Coil_147,Coil_148,Coil_149,Coil_150,Coil_151,Coil_152,Coil_153" & _
  ",Coil_154,Coil_155,Coil_156,Coil_157,Coil_158,Coil_159,Coil_160" & _
  ",Coil_161,Coil_162,Coil_163,Coil_164,Coil_165,Coil_166,Coil_167" & _
  ",Coil_168,Coil_169,Coil_170,Coil_171,Coil_172,Coil_173,Coil_174" & _
  ",Coil_175,Coil_176,Coil_177,Coil_178,Coil_179,Coil_180,Coil_181" & _
  ",Coil_182,Coil_183,Coil_184,Coil_185,Coil_186,Coil_187,Coil_188" & _
  ",Coil_189,Coil_190,Coil_191,Coil_192,Coil_193,Coil_194,Coil_195" & _
  ",Coil_196,Coil_197,Coil_198,Coil_199,Coil_200,Coil_201,Coil_202" & _
  ",Coil_203,Coil_204,Coil_205,Coil_206,Coil_207,Coil_208,Coil_209" & _
  ",Coil_210,Coil_211,Coil_212,Coil_213,Coil_214,Coil_215,Coil_216" & _
  ",Coil_217,Coil_218,Coil_219,Coil_220,Coil_221,Coil_222,Coil_223" & _
  ",Coil_224,Coil_225,Coil_226,Coil_227,Coil_228,Coil_229,Coil_230" & _
  ",Coil_231,Coil_232,Coil_233,Coil_234,Coil_235,Coil_236,Coil_237" & _
  ",Coil_238,Coil_239,Coil_240,Coil_241,Coil_242,Coil_243,Coil_244" & _
  ",Coil_245,Coil_246,Coil_247,Coil_248,Coil_249,Coil_250,Coil_251" & _
  ",Coil_252,Coil_253,Coil_254,Coil_255,Coil_256,Coil_257,Coil_258" & _
  ",Coil_259,Coil_260,Coil_261,Coil_262,Coil_263,Coil_264,Coil_265" & _
  ",Coil_266,Coil_267,Coil_268,Coil_269,Coil_270,Coil_271,Coil_272" & _
  ",Coil_273,Coil_274,Coil_275,Coil_276,Coil_277,Coil_278,Coil_279" & _
  ",Coil_280,Coil_281,Coil_282,Coil_283,Coil_284,Coil_285,Coil_286" & _
  ",Coil_287,Coil_288,Coil_289,Coil_290,Coil_291,Coil_292,Coil_293" & _
  ",Coil_294,Coil_295,Coil_296,Coil_297,Coil_298,Coil_299"), Array(_
  "NAME:SubtractParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  true)
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "StatorSlotInsu", _
  "Tool Parts:=", "Stator"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", true)
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "5398mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "1800mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "1180mm"), Array("NAME:Pair", "Name:=", _
  "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", _
  "2"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array(_
  "NAME:Pair", "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "5mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "SlotPitch", "Value:=", "1deg"), Array("NAME:Pair", _
  "Name:=", "CenterPitch", "Value:=", "2deg"), Array("NAME:Pair", "Name:=", _
  "Poles", "Value:=", "32"), Array("NAME:Pair", "Name:=", "WidthShoe", _
  "Value:=", "355mm"), Array("NAME:Pair", "Name:=", "HeightShoe", "Value:=", _
  "40mm"), Array("NAME:Pair", "Name:=", "WidthBody", "Value:=", "240mm"), _
  Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", "300mm"), Array(_
  "NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Offset", "Value:=", "10mm"), Array("NAME:Pair", "Name:=", _
  "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Off2_y", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", "Value:=", _
  "41.5mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", "1"), _
  Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "2423.705951218544mm"), Array("NAME:Pair", "Name:=", "InfoCore", _
  "Value:=", "0"))), Array("NAME:Attributes", "Name:=", "Rotor", "Flags:=", _
  "", "Color:=", "(132 132 193)", "Transparency:=", 0, _
  "PartCoordinateSystem:=", "Global", "MaterialName:=", "M22_24G_3DSF0.960", _
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
  "Name:=", "Length", "Value:=", "1180mm"), Array("NAME:Pair", "Name:=", _
  "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", _
  "2"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array(_
  "NAME:Pair", "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "5mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "SlotPitch", "Value:=", "1deg"), Array("NAME:Pair", _
  "Name:=", "CenterPitch", "Value:=", "2deg"), Array("NAME:Pair", "Name:=", _
  "Poles", "Value:=", "32"), Array("NAME:Pair", "Name:=", "WidthShoe", _
  "Value:=", "355mm"), Array("NAME:Pair", "Name:=", "HeightShoe", "Value:=", _
  "40mm"), Array("NAME:Pair", "Name:=", "WidthBody", "Value:=", "240mm"), _
  Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", "300mm"), Array(_
  "NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Offset", "Value:=", "10mm"), Array("NAME:Pair", "Name:=", _
  "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Off2_y", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", "Value:=", _
  "41.5mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", "1"), _
  Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "2423.705951218544mm"), Array("NAME:Pair", "Name:=", "InfoCore", _
  "Value:=", "2"))), Array("NAME:Attributes", "Name:=", "Field", "Flags:=", _
  "", "Color:=", "(255 128 192)", "Transparency:=", 0, _
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
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "5398mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "1800mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "1180mm"), Array("NAME:Pair", "Name:=", _
  "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", _
  "2"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array(_
  "NAME:Pair", "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "5mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "SlotPitch", "Value:=", "1deg"), Array("NAME:Pair", _
  "Name:=", "CenterPitch", "Value:=", "2deg"), Array("NAME:Pair", "Name:=", _
  "Poles", "Value:=", "32"), Array("NAME:Pair", "Name:=", "WidthShoe", _
  "Value:=", "355mm"), Array("NAME:Pair", "Name:=", "HeightShoe", "Value:=", _
  "40mm"), Array("NAME:Pair", "Name:=", "WidthBody", "Value:=", "240mm"), _
  Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", "300mm"), Array(_
  "NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Offset", "Value:=", "10mm"), Array("NAME:Pair", "Name:=", _
  "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Off2_y", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", "Value:=", _
  "41.5mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", "1"), _
  Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "2423.705951218544mm"), Array("NAME:Pair", "Name:=", "InfoCore", _
  "Value:=", "3"))), Array("NAME:Attributes", "Name:=", "Bar", "Flags:=", "", _
  "Color:=", "(119 198 206)", "Transparency:=", 0, "PartCoordinateSystem:=", _
  "Global", "MaterialName:=", "copper_100C", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Bar:CreateUserDefinedPart:1", _
  "LenRegion", "1180mm + 2*endRegion"
On Error Goto 0
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Bar", _
  "Tool Parts:=", "Tool"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", true)
oEditor.SeparateBody  Array("NAME:Selections", "Selections:=", "Bar")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.SetEddyEffect Array("NAME:Eddy Effect Setting", Array(_
  "NAME:EddyEffectVector", Array("NAME:Data", "Object Name:=", "Bar", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", "Bar_Separate1", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", "Bar_Separate2", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", "Bar_Separate3", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", "Bar_Separate4", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", "Bar_Separate5", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", "Bar_Separate6", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", "Bar_Separate7", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", "Bar_Separate8", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", "Bar_Separate9", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", _
  "Bar_Separate10", "Eddy Effect:=", true), Array("NAME:Data", _
  "Object Name:=", "Bar_Separate11", "Eddy Effect:=", true), Array(_
  "NAME:Data", "Object Name:=", "Bar_Separate12", "Eddy Effect:=", true), _
  Array("NAME:Data", "Object Name:=", "Bar_Separate13", "Eddy Effect:=", _
  true), Array("NAME:Data", "Object Name:=", "Bar_Separate14", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", _
  "Bar_Separate15", "Eddy Effect:=", true), Array("NAME:Data", _
  "Object Name:=", "Bar_Separate16", "Eddy Effect:=", true), Array(_
  "NAME:Data", "Object Name:=", "Bar_Separate17", "Eddy Effect:=", true), _
  Array("NAME:Data", "Object Name:=", "Bar_Separate18", "Eddy Effect:=", _
  true), Array("NAME:Data", "Object Name:=", "Bar_Separate19", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", _
  "Bar_Separate20", "Eddy Effect:=", true), Array("NAME:Data", _
  "Object Name:=", "Bar_Separate21", "Eddy Effect:=", true), Array(_
  "NAME:Data", "Object Name:=", "Bar_Separate22", "Eddy Effect:=", true), _
  Array("NAME:Data", "Object Name:=", "Bar_Separate23", "Eddy Effect:=", _
  true), Array("NAME:Data", "Object Name:=", "Bar_Separate24", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", _
  "Bar_Separate25", "Eddy Effect:=", true), Array("NAME:Data", _
  "Object Name:=", "Bar_Separate26", "Eddy Effect:=", true), Array(_
  "NAME:Data", "Object Name:=", "Bar_Separate27", "Eddy Effect:=", true), _
  Array("NAME:Data", "Object Name:=", "Bar_Separate28", "Eddy Effect:=", _
  true), Array("NAME:Data", "Object Name:=", "Bar_Separate29", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", _
  "Bar_Separate30", "Eddy Effect:=", true), Array("NAME:Data", _
  "Object Name:=", "Bar_Separate31", "Eddy Effect:=", true)))
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "OuterRegion"), _
  Array("NAME:ChangedProps", Array("NAME:Transparent", "Value:=", 0.75))))
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Shaft,Stator,Coil_0" & _
  ",Coil_1,Coil_2,Coil_3,Coil_4,Coil_5,Coil_6,Coil_7,Coil_8,Coil_9,Coil_10" & _
  ",Coil_11,Coil_12,Coil_13,Coil_14,Coil_15,Coil_16,Coil_17,Coil_18" & _
  ",Coil_19,Coil_20,Coil_21,Coil_22,Coil_23,Coil_24,Coil_25,Coil_26" & _
  ",Coil_27,Coil_28,Coil_29,Coil_30,Coil_31,Coil_32,Coil_33,Coil_34" & _
  ",Coil_35,Coil_36,Coil_37,Coil_38,Coil_39,Coil_40,Coil_41,Coil_42" & _
  ",Coil_43,Coil_44,Coil_45,Coil_46,Coil_47,Coil_48,Coil_49,Coil_50" & _
  ",Coil_51,Coil_52,Coil_53,Coil_54,Coil_55,Coil_56,Coil_57,Coil_58" & _
  ",Coil_59,Coil_60,Coil_61,Coil_62,Coil_63,Coil_64,Coil_65,Coil_66" & _
  ",Coil_67,Coil_68,Coil_69,Coil_70,Coil_71,Coil_72,Coil_73,Coil_74" & _
  ",Coil_75,Coil_76,Coil_77,Coil_78,Coil_79,Coil_80,Coil_81,Coil_82" & _
  ",Coil_83,Coil_84,Coil_85,Coil_86,Coil_87,Coil_88,Coil_89,Coil_90" & _
  ",Coil_91,Coil_92,Coil_93,Coil_94,Coil_95,Coil_96,Coil_97,Coil_98" & _
  ",Coil_99,Coil_100,Coil_101,Coil_102,Coil_103,Coil_104,Coil_105,Coil_106" & _
  ",Coil_107,Coil_108,Coil_109,Coil_110,Coil_111,Coil_112,Coil_113" & _
  ",Coil_114,Coil_115,Coil_116,Coil_117,Coil_118,Coil_119,Coil_120" & _
  ",Coil_121,Coil_122,Coil_123,Coil_124,Coil_125,Coil_126,Coil_127" & _
  ",Coil_128,Coil_129,Coil_130,Coil_131,Coil_132,Coil_133,Coil_134" & _
  ",Coil_135,Coil_136,Coil_137,Coil_138,Coil_139,Coil_140,Coil_141" & _
  ",Coil_142,Coil_143,Coil_144,Coil_145,Coil_146,Coil_147,Coil_148" & _
  ",Coil_149,Coil_150,Coil_151,Coil_152,Coil_153,Coil_154,Coil_155" & _
  ",Coil_156,Coil_157,Coil_158,Coil_159,Coil_160,Coil_161,Coil_162" & _
  ",Coil_163,Coil_164,Coil_165,Coil_166,Coil_167,Coil_168,Coil_169" & _
  ",Coil_170,Coil_171,Coil_172,Coil_173,Coil_174,Coil_175,Coil_176" & _
  ",Coil_177,Coil_178,Coil_179,Coil_180,Coil_181,Coil_182,Coil_183" & _
  ",Coil_184,Coil_185,Coil_186,Coil_187,Coil_188,Coil_189,Coil_190" & _
  ",Coil_191,Coil_192,Coil_193,Coil_194,Coil_195,Coil_196,Coil_197" & _
  ",Coil_198,Coil_199,Coil_200,Coil_201,Coil_202,Coil_203,Coil_204" & _
  ",Coil_205,Coil_206,Coil_207,Coil_208,Coil_209,Coil_210,Coil_211" & _
  ",Coil_212,Coil_213,Coil_214,Coil_215,Coil_216,Coil_217,Coil_218" & _
  ",Coil_219,Coil_220,Coil_221,Coil_222,Coil_223,Coil_224,Coil_225" & _
  ",Coil_226,Coil_227,Coil_228,Coil_229,Coil_230,Coil_231,Coil_232" & _
  ",Coil_233,Coil_234,Coil_235,Coil_236,Coil_237,Coil_238,Coil_239" & _
  ",Coil_240,Coil_241,Coil_242,Coil_243,Coil_244,Coil_245,Coil_246" & _
  ",Coil_247,Coil_248,Coil_249,Coil_250,Coil_251,Coil_252,Coil_253" & _
  ",Coil_254,Coil_255,Coil_256,Coil_257,Coil_258,Coil_259,Coil_260" & _
  ",Coil_261,Coil_262,Coil_263,Coil_264,Coil_265,Coil_266,Coil_267" & _
  ",Coil_268,Coil_269,Coil_270,Coil_271,Coil_272,Coil_273,Coil_274" & _
  ",Coil_275,Coil_276,Coil_277,Coil_278,Coil_279,Coil_280,Coil_281" & _
  ",Coil_282,Coil_283,Coil_284,Coil_285,Coil_286,Coil_287,Coil_288" & _
  ",Coil_289,Coil_290,Coil_291,Coil_292,Coil_293,Coil_294,Coil_295" & _
  ",Coil_296,Coil_297,Coil_298,Coil_299,Rotor,Field_0,Field_1,Field_2" & _
  ",Field_3,Field_4,Field_5,Field_6,Field_7,Field_8,Field_9,Field_10" & _
  ",Field_11,Field_12,Field_13,Field_14,Field_15,Field_16,Field_17" & _
  ",Field_18,Field_19,Field_20,Field_21,Field_22,Field_23,Field_24" & _
  ",Field_25,Field_26,Field_27,Field_28,Field_29,Field_30,Field_31,Bar" & _
  ",Bar_Separate1,Bar_Separate2,Bar_Separate3,Bar_Separate4,Bar_Separate5" & _
  ",Bar_Separate6,Bar_Separate7,Bar_Separate8,Bar_Separate9,Bar_Separate10" & _
  ",Bar_Separate11,Bar_Separate12,Bar_Separate13,Bar_Separate14" & _
  ",Bar_Separate15,Bar_Separate16,Bar_Separate17,Bar_Separate18" & _
  ",Bar_Separate19,Bar_Separate20,Bar_Separate21,Bar_Separate22" & _
  ",Bar_Separate23,Bar_Separate24,Bar_Separate25,Bar_Separate26" & _
  ",Bar_Separate27,Bar_Separate28,Bar_Separate29,Bar_Separate30" & _
  ",Bar_Separate31", "Tool Parts:=", "Tool"), Array(_
  "NAME:SubtractParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.FitAll
Set oModule = oDesign.GetModule("AnalysisSetup")
oModule.InsertSetup "Transient", Array("NAME:Setup1", "StopTime:=", "0.2s", _
  "TimeStep:=", "0.0002s")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.SetMinimumTimeStep "0.002ms" 
set oUnclassified = oEditor.GetObjectsInGroup("Unclassified")
Dim oObject
For Each oObject in oUnclassified
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", oObject), Array(_
  "NAME:ChangedProps", Array("NAME:Model", "Value:=", false))))
Next
