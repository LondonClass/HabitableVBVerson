VERSION 5.00
Begin VB.Form FormMain 
   AutoRedraw      =   -1  'True
   Caption         =   "宜居"
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8565
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   373
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   571
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4680
      Top             =   2400
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==枚举类型声明==

Private Enum FormEnum
    FormClosed = -1
    FormWhite = 0
    FormMainMenu = 1
    FormStartGame = 2
    FormGame = 3
    FormSettings = 4
End Enum

Private Enum EffectEnum
    EffectNone
    EffectOxygen
    EffectGas
    EffectWater
    EffectTempreture
    EffectReflectivity
    EffectHousing
    EffectResource
    EffectResearchPoint
    EffectPrestige
    EffectStorage
    EffectSolarPower
    EffectRunOff
End Enum

Private Enum ResourceEnum
    ResourceNone
    ResourceRock
    ResourceMineral
    ResourceMetel
    ResourceComposites
    ResourceFood
    ResourceCleanWater
    ResourceRobot
    ResourceFurniture
    ResourceElectricAppliance
    ResourceBioMaterial
    ResourceCount
End Enum

Private Enum ShowUI
    ShowUINone
    ShowUIFinance
    ShowUIPopulation
    ShowUIResearch
    ShowUIPlanet
    ShowUIResource
    ShowUILoadGame
    ShowUINewGame
    ShowUINewCampaign
    ShowUITutorial
End Enum

'在太空中的位置类型
Private Enum SpacePositionTypeEnum
    SpacePositionTypeLand '着陆
    SpacePositionTypeSurround '环绕
    SpacePositionTypeNavigation '航行
End Enum

Private Enum StarTypeEnum
    StarTypeNone
    StarTypeProto
    StarTypeMainSequence
    StarTypeDrawf
    StarTypeNeutron
    StarTypeBlackHole
End Enum

Private Enum MaterialStateEnum
    MaterialStateSolid
    MaterialStateLiquid
    MaterialStateGas
End Enum

Private Enum ModuleGroup
    ModuleGroupProduction
    ModuleGroupEnvironment
    ModuleGroupCount
End Enum

'==自定义类型声明==

Private Type PointApi
    X As Long
    Y As Long
End Type

Private Type Rectangle 'UI矩形
    Left As Long '矩形左边缘到容器边缘的距离
    Top As Long '矩形上边缘到容器边缘的距离
    Width As Long '矩形宽度,等于右边缘横坐标减左边缘横坐标
    Height As Long '矩形高度,等于下边缘纵坐标减上边缘纵坐标
End Type

Private Type UIObject 'UI对象，包含其类型、位置与大小、对鼠标的响应方式的信息
'    UIType As String 'UI对象的类型,决定了这个UI对象是按钮、图像、面板文字还是其他。决定了其绘制方式
'    'UIType的值有"button"：按钮,"clicker"：透明按钮等
'    Info As String '决定了UI的内容
    Position As Rectangle 'UI对象的大小与位置
    RealPosition As Rectangle 'UI对象的大小与位置
    Parent As Long 'UI的父对象ID
    Button As Long 'UI对象可以响应的鼠标按键
    ClickAction As String '点击UI对象所执行的指令
    Tooltip As String 'UI的鼠标悬停提示
End Type

Private Type Resource
    Type As ResourceEnum
    Amont As Double
End Type

Private Type Material
    Type As String '类型
    Mass As Double
End Type

Private Type MaterialType
    Name As String '名称
    State As MaterialStateEnum '物态
    LowTemp As Double
    LowTempTarget As String
    HighTemp As Double
    HighTempTarget As String
'    MolarMass As Double
End Type

Private Type Effect
    Type As EffectEnum
    Amont As Double
    EffectResources As Resource
End Type

'建筑类型
Private Type ModuleType
    Name As String
    Description As String
    Cost As Double
    Space As Double
    LivableRequire As Boolean
    Maintenance As Double
    Staff As Double
    Power As Double
    Resources() As Resource
    Effects() As Effect
    BuildTime As Double
    MaxTempreture As Double
    MaxPressure As Double
End Type

'建筑
Private Type Module
    Type As Long '类型
    Size As Long '建筑等级
    Construction As Double '建造进度
    Enabled As Boolean '是否启用
    EfficiencyModifier As Double '效率加成
    Efficiency As Double '实际效率
    Owner As Long '所有者
    Storage() As Resource '仓储
End Type

Private Type Colony
    Population As Double
    PopulationMoney As Double
    PopulationStorage() As Double
    PopulationReligion As Double
'    ReligiousUnity As Double
    Equality As Double
'    Immigration As Boolean
End Type

'恒星
Private Type Star
    Name As String
    Tag As String
    Type As StarTypeEnum
    Magnitude As Double
    Mass As Double
    Color As OLE_COLOR
End Type

Private Type Market
    Money() As Double
    Storage() As Double
    Prices() As Double
    Supply() As Double
    Demand() As Double
End Type

Private Type Planet
    Name As String
    Tag As String
    Radio As Double
    Tempreture As Double
    Reflectivity As Double
    RotationPeriod As Double '自转周期
    OrbitRadius As Double '公转半径
    Color As OLE_COLOR '行星的颜色
    
    OrbitRotation As Double
    
    Colonys As Colony
    Resources() As Double
    Transport() As Double
    Modules() As Module
    Materials() As Material
    
    Market As Market
    
    UtilizingBlock As Long
    BioMass As Double
    HomeWorld As Boolean '是否母星
    
    '运行中缓存
    Housing As Double '住房总量
    Storage As Double '总存储空间
End Type

'恒星系
Private Type System
    Name As String
    ID As Long
    Stars() As String
    Planets() As String
End Type

'在太空中的位置
Private Type SpacePosition
    Type As SpacePositionTypeEnum '位置类型
    Position1 As Double '航天器起点
    Position2 As Double '航天器终点
    Progress As Double '航天器距离起点的距离0-1
End Type

'宇宙飞船
Private Type Spacecraft
    Name As String '名字
    Maintenance As Double '维护费
    Power As Double '电量
    Position As SpacePosition '位置
    Population As Long '乘员数量
    Effects() As Effect '效果
    Owner As String '所有者
    Space As Double '存储空间
    Storage() As Resource '仓储物资
    Construction As Double '建造进度
    Enabled As Boolean
End Type

Private Type Technology
    Name As String
    NeedPoints As Long
    IsResearched As Boolean
End Type

Private Type FinancialInfo
    Funds As Double '拨款
    Contribution As Double '捐款
    Income As Double '总收入
    ColonyMaintence As Double '殖民地维护
    Salary As Double '工资
    Transport As Double '运输
    Expence As Double '总支出
    NetIncome As Double  '总计
End Type

Private Type GameEvent
    Title As String
    Content As String
    Options() As String
End Type

'==变量声明==

Dim FormOn As FormEnum '当前显示窗体
Dim FrequencyPerMillisecond As Double
Dim Fps As Long

'游戏内容变量
Dim GameDate As Date '游戏日期
Dim PreviousDate As Date '游戏日期的上一天
Dim Money As Double
Dim Prestige As Double
Dim ResearchPoint As Double

Dim System As System
Dim Planets() As Planet
Dim Stars() As Star
Dim ModuleTypes() As ModuleType
Dim MaterialTypes() As MaterialType
Dim Technologys() As Technology
Dim Spacecrafts() As Spacecraft
Dim Events() As GameEvent

Dim Funds As Long
Dim FinancialInfo As FinancialInfo
Dim PreviousFinancialInfo As FinancialInfo

'UI相关变量
Dim ShowEvents() As Long
Dim ShowingUI As ShowUI
Dim SelectMenuButton As Long '临时变量。记录UI界面的菜单栏
Dim SelectModule As Long
Dim SelectSpacecraft As Long
Dim SelectPlanet As Long
Dim DrawModuleOffsetTarget As Long
Dim DrawModuleOffset As Long
Dim IsShowPauseMenu As Boolean
Dim IsShowDebugWindow As Boolean
Dim IsWriteLog As Boolean

Dim GameSpeed As Long '游戏速度

Dim GameLog As String '程序生成的日志文件

'Dim MouseButton As Long
Dim UIObjectList() As UIObject '用于UI对鼠标的响应而创建的列表

'==api声明==

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long '计时器
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long '获取计时器频率
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As PointApi) As Long '获取光标位置

'向UI列表中添加UI，并返回其在UI列表中的序号
Private Function AddUIObjectList(Position As Rectangle, Optional ByVal ClickAction As String = "", Optional ByVal Button As Long = 1, Optional ByVal Tooltip As String = "", Optional ByVal Parent As Long = 0) As Long
    ReDim Preserve UIObjectList(UBound(UIObjectList) + 1)
    With UIObjectList(UBound(UIObjectList))
        .Position = Position
        .Parent = Parent
        .RealPosition = RectangleTranslation(Position, UIObjectList(Parent).RealPosition.Left, UIObjectList(Parent).RealPosition.Top)
        .Button = Button
        .ClickAction = ClickAction
        .Tooltip = Tooltip
    End With
    AddUIObjectList = UBound(UIObjectList)
End Function

'清除选择
Private Sub ClearSelect()
    SelectModule = 0
    SelectSpacecraft = 0
End Sub

Private Function CloseEvent(n As Long)
    ReDim ShowEvents(0)
End Function

'计算殖民地需求
'Private Function ColonyDemandCalculation(Colony As Colony, ResourceType As ResourceEnum) As Double
'    Colony.PopulationMoney*1
'    Select Case ResourceType
'    Case ResourceFood
'        ColonyDemandCalculation = 1
'    Case ResourceCleanWater
'        ColonyDemandCalculation = 1
'    Case Else
'        ColonyDemandCalculation = 0
'    End Select
'End Function

'每日计算
Private Sub DailyCaculate()
    Dim i As Long, j As Long
    
    PreviousDate = GameDate
    GameDate = GameDate + 1
    
    '行星计算
    For i = 1 To UBound(Planets)
        DailyCaculatePlanet Planets(i)
    Next
    
    '月度计算
    If Month(PreviousDate) <> Month(GameDate) Then
        '计算财务信息
        With FinancialInfo
            Money = Money + Funds
            .Funds = Funds
            Money = Money + Int(10 * Prestige ^ 0.5)
            .Contribution = Int(10 * Prestige ^ 0.5)
            .Income = .Funds + .Contribution
            For i = 1 To UBound(Planets)
                With Planets(i)
                    If Not .HomeWorld Then
                        '模块经济计算
                        For j = 1 To UBound(.Modules)
                            Money = Money + ModuleTypes(.Modules(j).Type).Maintenance * .Modules(j).Efficiency * .Modules(j).Size
                            FinancialInfo.ColonyMaintence = FinancialInfo.ColonyMaintence + ModuleTypes(.Modules(j).Type).Maintenance * .Modules(j).Efficiency * .Modules(j).Size
                        Next
                        '工资计算
                        Money = Money - 0.005 * .Colonys.Population
                        FinancialInfo.Salary = FinancialInfo.Salary - 0.005 * .Colonys.Population
                    End If
                End With
            Next
            .Expence = .ColonyMaintence + .Salary + .Transport
            .NetIncome = .Income + .Expence
        End With
        PreviousFinancialInfo = FinancialInfo
        FinancialInfo = FinancialInfoCreate
    End If
    
'    If Year(PreviousDate) <> Year(GameDate) Then
End Sub

'市场每日计算
Private Sub DailyCaculateMarket(Market As Market, Planet As Planet)
    Dim Demand As Double
    Dim ActualBought As Double
    Dim TempMoney As Double
    Dim TempGoods As Double
    Dim TempPrice As Double
    
    With Market
        .Money(ResourceFood) = .Money(ResourceFood) + 1 '临时提供流动性
        
        If .Storage(ResourceFood) > 1 And Planet.Colonys.PopulationMoney > 1 Then
            '计算交易
            .Prices(ResourceFood) = .Money(ResourceFood) / .Storage(ResourceFood)
            
            '计算人口需求
            Demand = Min(Planet.Colonys.Population, Planet.Colonys.PopulationMoney / .Prices(ResourceFood))
            
            '判断市场是否有足够物资
            If Demand < .Storage(ResourceFood) Then
                '计算市场应有金钱
                TempMoney = MarketGetProduct(Market, ResourceFood) / (.Storage(ResourceFood) - Demand)
                
                TempPrice = (TempMoney - .Money(ResourceFood)) / Demand
                
                If (TempMoney - .Money(ResourceFood)) < Planet.Colonys.PopulationMoney Then '检测钱是否够买
                    ActualBought = Demand
                    
                    Planet.Colonys.PopulationMoney = Planet.Colonys.PopulationMoney - ActualBought * TempPrice
                    .Money(ResourceFood) = .Money(ResourceFood) + ActualBought * TempPrice
                    .Storage(ResourceFood) = .Storage(ResourceFood) - ActualBought
                    Planet.Colonys.PopulationStorage(ResourceFood) = Planet.Colonys.PopulationStorage(ResourceFood) + ActualBought
                Else
                    '只买部分，将所有钱用来购买物资
                    TempGoods = MarketGetProduct(Market, ResourceFood) / (.Money(ResourceFood) + Planet.Colonys.PopulationMoney)
                    ActualBought = .Storage(ResourceFood) - TempGoods
                    TempPrice = Planet.Colonys.PopulationMoney / ActualBought
                    
                    Planet.Colonys.PopulationMoney = Planet.Colonys.PopulationMoney - ActualBought * TempPrice
                    .Money(ResourceFood) = .Money(ResourceFood) + ActualBought * TempPrice
                    .Storage(ResourceFood) = .Storage(ResourceFood) - ActualBought
                    Planet.Colonys.PopulationStorage(ResourceFood) = Planet.Colonys.PopulationStorage(ResourceFood) + ActualBought
                End If
            End If
        End If
    End With
End Sub

'模块每日计算
Private Sub DailyCaculateModule(Module As Module, Planet As Planet, ByVal ID As Long)
    With Module
        '建造模块
        If .Construction < 1 Then
            .Construction = .Construction + 1 / ModuleTypes(.Type).BuildTime
            If .Construction > 1 Then
                .Construction = 1
                .Size = .Size + 1
            End If
        End If
        
        '检测模块是否过热超压，并摧毁过热超压模块
        If ModuleTypes(.Type).MaxTempreture < Planet.Tempreture Then
            MsgBox Planet.Name & "上的建筑" & ModuleTypes(.Type).Name & "由于温度过高被毁"
            PlanetDeleteMoudle Planet, ID
        End If
        If ModuleTypes(.Type).MaxPressure < PlanetGetPressure(Planet) Then
            MsgBox Planet.Name & "上的建筑" & ModuleTypes(.Type).Name & "由于气压过大被毁"
            PlanetDeleteMoudle Planet, ID
        End If
        
        '计算模块效率
        MoudleEfficiencyCalculate Module, Planet, ID

        '计算模块效果
        MoudleEffectCalculate Module, Planet
    End With
End Sub

'行星每日计算
Private Sub DailyCaculatePlanet(Planet As Planet)
    Dim i As Long
    Dim ResourceSend As Double '临时变量，分配的资源数量
    
    With Planet
        .Housing = 0
        .Storage = 0
        
        '飞船效果计算
        For i = 1 To UBound(Spacecrafts)
            If SpacecraftGetPlanet(Spacecrafts(i)) = PlanetGetID(Planet) Then
                DailyCaculateSpacecraft Spacecrafts(i)
            End If
        Next
    
        '星球旋转
        .OrbitRotation = .OrbitRotation - 2 * 4 * Atn(1) / PlanetGetOrbitalPeriod(Stars(1), Planet)
        
        '温度计算
        .Tempreture = .Tempreture - 10 * (.Tempreture / 273.15) ^ 4 * (1 - PlanetGetGreenhouseEffect(Planet)) + 12 * (1 - GetReflectivity(Planet)) * (1.496 / .OrbitRadius) ^ 2
        
        '物态计算
        For i = UBound(.Materials) To 1 Step -1
            If .Materials(i).Mass <= 0 Then
                PlanetDeleteMaterial Planet, .Materials(i).Type
            End If
        Next
        For i = UBound(.Materials) To 1 Step -1
            If MaterialGetType(.Materials(i)).LowTempTarget <> "" Then
                If .Tempreture < MaterialGetType(.Materials(i)).LowTemp Then
                    PlanetChangeMaterial Planet, MaterialGetType(.Materials(i)).LowTempTarget, .Materials(i).Mass
                    PlanetDeleteMaterial Planet, .Materials(i).Type
                End If
            End If
        Next
        For i = UBound(.Materials) To 1 Step -1
            If MaterialGetType(.Materials(i)).HighTempTarget <> "" Then
                If .Tempreture > MaterialGetType(.Materials(i)).HighTemp Then
                    PlanetChangeMaterial Planet, MaterialGetType(.Materials(i)).HighTempTarget, .Materials(i).Mass
                    PlanetDeleteMaterial Planet, .Materials(i).Type
                End If
            End If
        Next
        
'        '清空市场数据
'        If .HomeWorld Then
'            '清空需求数据
'            For i = 1 To UBound(.MarketSupply)
'                .MarketSupply(i) = 0
'            Next
'        End If

        '模块计算
        For i = 1 To UBound(.Modules)
            DailyCaculateModule .Modules(i), Planet, i
        Next
        If .Tempreture < 0 Then .Tempreture = 0 '防止温度为负
        
        '资源的损耗
        For i = 1 To UBound(.Resources)
            If .Resources(i) > Planet.Storage Then
                .Resources(i) = .Resources(i) - 0.01 * (.Resources(i) - Planet.Storage)
            End If
        Next
    
        '持续运输资源
        For i = 1 To UBound(.Resources)
            If .Transport(i) > 0 Then
                If Money > .Transport(i) Then
                    Money = Money - .Transport(i)
                    FinancialInfo.Transport = FinancialInfo.Transport - .Transport(i)
                    .Resources(i) = .Resources(i) + .Transport(i)
                End If
            End If
            If .Transport(i) < 0 Then
                If .Resources(i) > -.Transport(i) Then
                    Money = Money - .Transport(i)
                    FinancialInfo.Transport = FinancialInfo.Transport - .Transport(i)
                    .Resources(i) = .Resources(i) + .Transport(i)
                End If
            End If
        Next
        
        '市场计算
        If .HomeWorld Then
            DailyCaculateMarket .Market, Planet
        End If
        
        '分配计算
'        If Not .HomeWorld Then
'            ResourceSend = Min(.Resources(ResourceFood), .Colonys.Population)
'            .Resources(ResourceFood) = .Resources(ResourceFood) - ResourceSend
'            .Colonys.PopulationStorage(ResourceFood) = .Colonys.PopulationStorage(ResourceFood) + ResourceSend
'
'            ResourceSend = Min(.Resources(ResourceCleanWater), .Colonys.Population)
'            .Resources(ResourceCleanWater) = .Resources(ResourceCleanWater) - ResourceSend
'            .Colonys.PopulationStorage(ResourceCleanWater) = .Colonys.PopulationStorage(ResourceCleanWater) + ResourceSend
'        End If
    
        '人口计算
        DailyCaculatePopulation .Colonys, Planet
    End With
End Sub

'人口每日计算
Private Sub DailyCaculatePopulation(Colony As Colony, Planet As Planet)
    With Colony
        .Population = .Population * (1 + GetGrowthRate(Planet))
        If .Population < 0 Then
            .Population = 0
        End If
        If .Population > 0 And .Population < 1000000# Then
            .Population = RandomInt(.Population)
        End If
    
        If Planet.Resources(ResourceFood) >= .Population * 0.0001 Then
            PlanetAddResource Planet, ResourceFood, -.Population * 0.0001
        Else
            Planet.Resources(ResourceFood) = 0
        End If
        
        If Planet.Resources(ResourceCleanWater) >= .Population * 0.0001 Then
            PlanetAddResource Planet, ResourceCleanWater, -.Population * 0.0001
        Else
            Planet.Resources(ResourceCleanWater) = 0
        End If
        
        .PopulationMoney = .PopulationMoney + .Population
    End With
End Sub

'飞船每日计算
Private Sub DailyCaculateSpacecraft(Spacecraft As Spacecraft)
    Dim i As Long
    
    With Spacecraft
        '建造飞船
        If .Construction < 1 Then
            .Construction = .Construction + 1 / 60
            If .Construction > 1 Then
                .Construction = 1
            End If
        End If

        '计算飞船效果
        If SpacecraftGetPlanet(Spacecraft) <> 0 Then
            For i = 1 To UBound(Spacecraft.Effects)
                EffectCalculate Spacecraft.Effects(i), Planets(SpacecraftGetPlanet(Spacecraft)), 1
            Next
        End If
    End With
End Sub

'删除飞行器
Private Sub DeleteSpacecraft(ID As Long)
    Dim i As Long
    For i = ID To UBound(Spacecrafts) - 1
        Spacecrafts(i) = Spacecrafts(i + 1)
    Next
    ReDim Preserve Spacecrafts(UBound(Spacecrafts) - 1)
End Sub

Public Sub DoActionSentence(ByVal Action As String) '解析指令语句,并执行
    Dim Words() As String
    Dim Name As String
    Dim Parameters() As Variant
    Dim i As Long
    
    '将指令列表按行分割，每一行都先设置缩进再打印
    Words = Split(Action, " ")
    
    Name = Words(LBound(Words))
    
    If UBound(Words) > 0 Then
        ReDim Parameters(UBound(Words) - 1)
        
        For i = LBound(Words) + 1 To UBound(Words)
            Parameters(i - 1) = Words(i)
        Next
    End If
    
    DoSystemAction Name, Parameters
End Sub

Private Sub DoSystemAction(ByVal Name As String, Parameters() As Variant) '鼠标点击相应
    Dim ID As Long
    
    Select Case Name
    Case "" '无效果
    Case "none" '无效果
    Case "develop_population"
'    Case 1 '发展人口 参数0为星球编号 参数1为人口数量
        If Money > 10 Then
            ID = FindPlanetWithTag(CStr(Parameters(0)))
            With Planets(ID)
                If .Colonys.Population < Planets(ID).Housing Then
                    Money = Money - Parameters(1)
                    .Colonys.Population = .Colonys.Population + Parameters(1)
                Else
                    MsgBox "居住空间不足"
                End If
            End With
        Else
            MsgBox "资金不足"
        End If
    Case "bulid_module"
'    Case 2 '建造建筑 参数0为星球编号 参数1为建筑种类
        PlanetBuildModule Planets(FindPlanetWithTag(CStr(Parameters(0)))), Parameters(1)
    Case "change_showing_ui"
'    Case 3 '切换正在显示的UI 参数0为切换到的UI编号
        ShowingUI = Parameters(0)
        SelectMenuButton = 0
    Case "set_select_module"
'    Case 4 '选择建筑 参数0为建筑编号
        SelectModule = Parameters(0)
    Case "switch_select_module" '切换选中的建筑
        If SelectModule <> Parameters(0) Then
            ClearSelect
            SelectModule = Parameters(0)
        Else
            SelectModule = 0
        End If
    Case "clear_select"
        ClearSelect
    Case "set_select_spacecraft"
        SelectSpacecraft = Parameters(0)
    Case "switch_select_spacecraft" '切换选中的飞船
        If SelectSpacecraft <> Parameters(0) Then
            ClearSelect
            SelectSpacecraft = Parameters(0)
        Else
            SelectSpacecraft = 0
        End If
'    Case 5 '建筑翻页 参数0为翻页距离
    Case "change_module_offset"
        DrawModuleOffsetTarget = DrawModuleOffsetTarget + Parameters(0)
        If DrawModuleOffsetTarget < 0 Then DrawModuleOffsetTarget = 0
    Case "dismantle_module"
'    Case 6 '拆除建筑 参数0为星球编号 参数1为建筑编号
        PlanetDeleteMoudle Planets(FindPlanetWithTag(CStr(Parameters(0)))), CLng(Parameters(1))
    Case "dismantle_spacecraft"
        DeleteSpacecraft CLng(Parameters(0))
    Case "transport_resources"
'    Case 7 '一次性运输资源 参数0为星球编号 参数1为资源编号 参数2为数量
        If Money > 10 Then
            With Planets(FindPlanetWithTag(CStr(Parameters(0))))
                If UBound(.Resources) >= Parameters(1) Then
                    Money = Money - 10
                    .Resources(Parameters(1)) = .Resources(Parameters(1)) + Parameters(2)
                Else
                    MsgBox "下标越界"
                End If
            End With
        Else
            MsgBox "资金不足"
        End If
    Case "add_transport_resources"
'    Case 8 '增加持续运输资源 参数0为星球编号 参数1为资源编号 参数2为数量
'    Case 9 '减少持续运输资源
        With Planets(FindPlanetWithTag(CStr(Parameters(0))))
            If UBound(.Resources) >= Parameters(1) Then
                .Transport(Parameters(1)) = .Transport(Parameters(1)) + Parameters(2)
            Else
                MsgBox "运输资源错误：下标越界"
            End If
        End With
    Case "change_window"
'        Case 10 '切换窗口 参数0为切换到的窗口编号
        FormOn = Parameters(0)
        If Parameters(0) = FormGame Then
            GameInitialization
        End If
    Case "add_utilizing_block"
'    Case 11 '增加建筑空间 参数0为添加建筑空间的星球编号
        ID = FindPlanetWithTag(CStr(Parameters(0)))
        With Planets(ID)
            If Money > GetBlockCost(Planets(ID)) Then
                Money = Money - GetBlockCost(Planets(ID))
                .UtilizingBlock = .UtilizingBlock + 1
            Else
                MsgBox "资金不足"
            End If
        End With
    Case "select_planet"
'    Case 12 '选择星球
        SelectPlanet = Parameters(0)
        ShowingUI = ShowUINone
    Case "close_event"
'    Case 13 '关闭事件 参数0为时间编号
        CloseEvent CLng(Parameters(0))
'    Case 14 '改变游戏速度
    Case "set_speed"
        GameSpeed = Parameters(0)
    Case "swith_pause_menu"
'    Case 15 '显示/隐藏暂停界面
        IsShowPauseMenu = Not IsShowPauseMenu
    Case "switch_module_enabled"
'    Case 16 '启/禁用建筑 参数0为星球编号 参数1为建筑编号
        ID = FindPlanetWithTag(CStr(Parameters(0)))
        Planets(ID).Modules(Parameters(1)).Enabled = Not Planets(ID).Modules(Parameters(1)).Enabled
        
    Case "switch_spacecraft_enabled" '启/禁用飞船 参数0为飞船编号
        Spacecrafts(Parameters(0)).Enabled = Not Spacecrafts(Parameters(0)).Enabled
    
    Case "move_spacecraft_to" '启/禁用飞船 参数0为飞船编号 参数1为星球编号
        Spacecrafts(Parameters(0)).Position.Position1 = Parameters(1)
    
    Case "unload_spacecraft_storage" '卸载飞船上的物资 参数0为飞船编号
        SpacecraftUnloadStorageAll Spacecrafts(Parameters(0))
    
    Case "expand_module"
'    Case 17 '扩建建筑 参数0为星球编号 参数1为建筑编号
        ExpansionModule Planets(FindPlanetWithTag(CStr(Parameters(0)))), Parameters(1)
    Case "set_menu_button"
'    Case 18 '改变界面按钮
        SelectMenuButton = Parameters(0)
'    Case 19 '点击仓库资源
    Case "save_game" '保存游戏
        SaveGame
    Case "msgbox" '保存游戏
        MsgBox Parameters(0)
    End Select
End Sub

'绘制底部菜单
Private Sub DrawBottomBar(Position As Rectangle)
    With Position
        DrawButtonWithUI Position, "财务", "system(""change_showing_ui " & ShowUIFinance & """)"
        .Left = .Left + .Width + 10
        
        DrawButtonWithUI Position, "人口", "system(""change_showing_ui " & ShowUIPopulation & """)"
        .Left = .Left + .Width + 10
        
        DrawButtonWithUI Position, "研究", "system(""change_showing_ui " & ShowUIResearch & """)"
        .Left = .Left + .Width + 10
        
        If SelectPlanet > 0 Then
            DrawButtonWithUI Position, "星球", "system(""change_showing_ui " & ShowUIPlanet & """)"
            .Left = .Left + .Width + 10
            
            DrawButtonWithUI Position, "资源", "system(""change_showing_ui " & ShowUIResource & """)"
        End If
    End With
End Sub

'绘制标准按钮
Private Sub DrawButton(Position As Rectangle, ByVal Text As String, Optional ByVal Color As OLE_COLOR = &HC0C0C0)
    RectangleDraw Position, Color '绘制
    MyPrint Text, Position.Left + 0.5 * Position.Width, Position.Top + 0.5 * Position.Height, 1 '绘制文字
End Sub

'绘制标准按钮并添加UI
Private Sub DrawButtonWithUI(Position As Rectangle, Optional ByVal Text As String = "NO_TEXT", Optional ByVal ClickAction As String = "", Optional ByVal Tooltip As String = "", Optional ByVal Color As OLE_COLOR = &HC0C0C0, Optional ByVal Parent As Long = 0, Optional ByVal Button As Long = 1)
    DrawButton Position, Text, Color
    AddUIObjectList Position, ClickAction, Button, Tooltip, Parent
End Sub

Private Sub DrawCloseButton(ByVal X As Long, ByVal Y As Long, Optional ByVal ClickAction As String = "system(""change_showing_ui " & ShowUINone & """)", Optional ByVal Size As Long = 15) '绘制关闭按钮
    DrawButtonWithUI RectangleCreate(X, Y, Size, Size), "×", ClickAction, , RGB(255, 64, 64)
End Sub

Private Sub DrawDebugWindow() '绘制debug窗口
    Dim ButtonPosition As Rectangle
    
    FillColor = vbWhite
    Line (50, 50)-(300, GetCenterY + 100), , B
End Sub

Private Sub DrawEvent(ByVal X As Long, ByVal Y As Long, ByVal ID As Long) '绘制事件界面
    Dim i As Long
    Dim ButtonPosition As Rectangle
    Dim PrintY As Long

    ForeColor = vbBlack
    FillStyle = 0
    FillColor = vbWhite
    FontSize = 11
    Line (X, Y)-(X + 300, Y + 200), , B
    FillColor = RGB(160, 200, 255)
    Line (X, Y)-(X + 300, Y + TextHeight(Events(ID).Title) + 10), , B
    MyPrint Events(ID).Title, X + 150 - TextWidth(Events(ID).Title) * 0.5, Y + 5

    FontSize = 9
    PrintY = Y + TextHeight(Events(ID).Title) + 10
    MyPrint Events(ID).Content, X + 5, PrintY + 5

    DrawButtonWithUI RectangleCreate(X + 20, Y + 160, 260, 30), Events(ID).Options(1), "system(""close_event " & ID & """)"
End Sub

Private Sub DrawGame() '绘制游戏界面到屏幕
    '绘制游戏主体部分
    If SelectPlanet = 0 Then
        SystemDraw System
    Else
        DrawModules Planets(SelectPlanet)
        DrawSpacecrafts Planets(SelectPlanet)
    End If
    '绘制游戏UI
    DrawGameUI
End Sub

Private Sub DrawGameUI()
    Dim i As Long
    Dim ButtonPosition As Rectangle
    
    CurrentX = 0
    CurrentY = 0
    Print "FPS:" & Fps
    Print "日期:" & GameDate
    Print "资金:" & Int(Money)
    Print "声望:" & Int(Prestige)
    Print "研究点数:" & Int(ResearchPoint)
    
    '显示选中模块效果
    If SelectPlanet > 0 Then
        If SelectModule > 0 And SelectModule <= UBound(Planets(SelectPlanet).Modules) Then
            ModuleDrawUI Planets(SelectPlanet).Modules(SelectModule), 500, 10
        End If
    End If
    
    '显示选中航天器效果
    If SelectSpacecraft > 0 And SelectSpacecraft <= UBound(Spacecrafts) Then
        SpacecraftDrawUI Spacecrafts(SelectSpacecraft), 500, 10
    End If
    
    DrawSpeedControlBar RectangleCreate(ScaleWidth - 170, 20, 30, 20) '绘制速度控制器。位于屏幕右上角
    
    '绘制打开的UI窗口
    If ShowingUI <> ShowUINone Then
        Select Case ShowingUI
        Case ShowUIFinance
            DrawGameUIFinance 340, 20
        Case ShowUIPopulation
            DrawGameUIPopulation 340, 20
        Case ShowUIResearch
            DrawUIResearch 340, 20
        Case ShowUIPlanet
            DrawGameUIPlanet Planets(SelectPlanet), 340, 20
        Case ShowUIResource
            DrawGameUIResource Planets(SelectPlanet), 340, 20
        End Select
    End If
    
    DrawBottomBar RectangleCreate(30, ScaleHeight - 70, 70, 40) '绘制底部菜单
    
    If SelectPlanet > 0 Then DrawButtonWithUI RectangleCreate(ScaleWidth - 50, 50, 30, 30), "星系", "system(""select_planet 0"")" '绘制返回星系按钮
    
    DrawToolTips '绘制提示
    
    '显示事件
    For i = 1 To UBound(ShowEvents)
        DrawEvent GetCenterX - 150, GetCenterY - 100, ShowEvents(i)
    Next
     
    If IsShowPauseMenu Then DrawPauseMenu
End Sub

Private Sub DrawGameUIFinance(ByVal X As Long, ByVal Y As Long)
    Dim DrawPosition As Rectangle
    
    FillColor = vbWhite
    Line (X, Y)-(X + 400, Y + 300), , B
    
    MyPrint "财务", X + 10, Y + 15
    
    With PreviousFinancialInfo
        CurrentX = X + 10
        Print "拨款:" & Format(.Funds, "0.00")
        CurrentX = X + 10
        Print "捐款:" & Format(.Contribution, "0.00")
        CurrentX = X + 10
        Print "总收入:" & Format(.Income, "0.00")
        CurrentX = X + 10
        Print "殖民地维护:" & Format(.ColonyMaintence, "0.00")
        CurrentX = X + 10
        Print "工资:" & Format(.Salary, "0.00")
        CurrentX = X + 10
        Print "运输:" & Format(.Transport, "0.00")
        CurrentX = X + 10
        Print "总支出:" & Format(.Expence, "0.00")
        CurrentX = X + 10
        Print "总计:" & Format(.NetIncome, "0.00")
    End With
    
    DrawCloseButton X + 330, Y + 5
End Sub

Private Sub DrawUILoadGame(ByVal X As Long, ByVal Y As Long)
    FillColor = vbWhite
    Line (X, Y)-(X + 400, Y + 300), , B
    
    MyPrint "没有存档", X + 10, Y + 15
End Sub

Private Sub DrawUINewCampaign(ByVal X As Long, ByVal Y As Long)
    FillColor = vbWhite
    Line (X, Y)-(X + 400, Y + 300), , B
    
    CurrentX = X + 10
    CurrentY = Y + 15
    Print "没有场景"
End Sub

Private Sub DrawUINewGame(ByVal X As Long, ByVal Y As Long)
    FillColor = vbWhite
    Line (X, Y)-(X + 400, Y + 300), , B
    
    CurrentX = X + 10
    CurrentY = Y + 15
    Print "新游戏"
End Sub

Private Sub DrawGameUIPlanet(Planet As Planet, ByVal X As Long, ByVal Y As Long)
    Dim DrawPosition As Rectangle
    Dim i As Long
    
    FillColor = vbWhite
    Line (X, Y)-(X + 400, Y + 300), , B
    
    CurrentX = X + 10
    CurrentY = Y + 15
    Print "星球"
    
    If SelectPlanet = 0 Then Exit Sub
    
    Select Case SelectMenuButton
    Case 0
        DrawButtonWithUI RectangleCreate(X + 10, Y + 40, 60, 20), "基本信息", "system(""set_menu_button 0"")", , RGB(128, 128, 128)
        DrawButtonWithUI RectangleCreate(X + 80, Y + 40, 40, 20), "文明", "system(""set_menu_button 1"")"
        DrawButtonWithUI RectangleCreate(X + 130, Y + 40, 40, 20), "市场", "system(""set_menu_button 2"")"
    
        With Planet
            CurrentX = X + 10
            CurrentY = Y + 70
            Print .Name
            
            CurrentX = X + 10
            Print "固态组成:"
            If PlanetGetMass(Planet, MaterialStateSolid) <> 0 Then
                With Planet
                    For i = 1 To UBound(.Materials)
                        With .Materials(i)
                            If MaterialGetState(Planet.Materials(i)) = MaterialStateSolid Then
                                CurrentX = X + 10
                                Print .Type & ":" & Format(PlanetGetMass(Planet, , .Type) / PlanetGetMass(Planet, MaterialStateSolid), "0.00%")
                            End If
                        End With
                    Next
                End With
            Else
                CurrentX = X + 10
                Print "无"
            End If
            
            CurrentX = X + 10
            Print "海洋组成:"
            If PlanetGetMass(Planet, MaterialStateLiquid) <> 0 Then
                With Planet
                    For i = 1 To UBound(.Materials)
                        With .Materials(i)
                            If MaterialGetState(Planet.Materials(i)) = MaterialStateLiquid Then
                                CurrentX = X + 10
                                Print .Type & ":" & Format(PlanetGetMass(Planet, , .Type) / PlanetGetMass(Planet, MaterialStateLiquid), "0.00%")
                            End If
                        End With
                    Next
                End With
            Else
                CurrentX = X + 10
                Print "无"
            End If
            
            CurrentX = X + 10
            Print "大气组成:"
            If PlanetGetMass(Planet, MaterialStateGas) <> 0 Then
                With Planet
                    For i = 1 To UBound(.Materials)
                        With .Materials(i)
                            If MaterialGetState(Planet.Materials(i)) = MaterialStateGas Then
                                CurrentX = X + 10
                                Print .Type & ":" & Format(PlanetGetMass(Planet, , .Type) / PlanetGetMass(Planet, MaterialStateGas), "0.00%")
                            End If
                        End With
                    Next
                End With
            Else
                CurrentX = X + 10
                Print "无"
            End If
            
            CurrentX = X + 10
            Print "氧气:" & Format(PlanetGetOxygen(Planet), "0.000000%");
            CurrentX = X + 140
            If PlanetGetOxygen(Planet) < 0.1 Then
                Print "氧气过少"
            Else
                If PlanetGetOxygen(Planet) <= 0.32 Then
                    If PlanetGetOxygen(Planet) >= 0.18 And PlanetGetOxygen(Planet) <= 0.24 Then
                        Print "氧气适宜居住"
                    Else
                        Print "氧气适宜植物生命"
                    End If
                Else
                    Print "氧气过多"
                End If
            End If
            CurrentX = X + 10
            Print "水:" & Format(PlanetGetWater(Planet), "0.000000%");
            CurrentX = X + 140
            If PlanetGetWater(Planet) < 0.1 Then
                Print "水分过少"
            Else
                If PlanetGetWater(Planet) <= 0.9 Then
                    If PlanetGetWater(Planet) >= 0.25 And PlanetGetWater(Planet) <= 0.75 Then
                        Print "水分适宜居住"
                    Else
                        Print "水分适宜植物生命"
                    End If
                Else
                    Print "水分过多"
                End If
            End If
            CurrentX = X + 10
            Print "反射率" & Format(GetReflectivity(Planet), "0.00%")
            CurrentX = X + 10
            Print "温度:" & Format(.Tempreture - 273.15, "0.00") & "℃";
            CurrentX = X + 140
            If Planet.Tempreture < 200 Then
                Print "温度过低"
            Else
                If Planet.Tempreture <= 374 Then
                    If Planet.Tempreture >= 237 And Planet.Tempreture <= 337 Then
                        Print "温度适宜居住"
                    Else
                        Print "温度适宜植物生命"
                    End If
                Else
                    Print "温度过高"
                End If
            End If
            CurrentX = X + 10
            Print "气压:" & Format(PlanetGetPressure(Planet) / 1000, "0.00") & "kPa";
            CurrentX = X + 140
            If PlanetGetPressure(Planet) < 10000 Then
                Print "气压过低"
            Else
                If PlanetGetPressure(Planet) <= 190000 Then
                    If PlanetGetPressure(Planet) >= 50000 And PlanetGetPressure(Planet) <= 150000 Then
                        Print "气压适宜居住"
                    Else
                        Print "气压适宜植物生命"
                    End If
                Else
                    Print "气压过高"
                End If
            End If
            CurrentX = X + 10
            Print "温室效应:" & Format(PlanetGetGreenhouseEffect(Planet), "0.00%")
            CurrentX = X + 10
            Print "公转周期:" & Format(PlanetGetOrbitalPeriod(Stars(1), Planet), "0.00") & "天"
            CurrentX = X + 10
            Print "自转周期:" & Format(.RotationPeriod, "0.00") & "天"
            CurrentX = X + 10
            Print "太阳能:" & Format(PlanetGetSolarPower(Stars(1), Planet), "0.00%")
        End With
        
        CurrentX = X + 10
        Print "电力产出:" & GetPowerProduce(Planet)
        CurrentX = X + 10
        Print "电力消耗:" & GetPowerUse(Planet)
        CurrentX = X + 10
        Print "电力充足率:" & Format(GetPowerAdyquacy(Planet), "0.00%")
        
        CurrentX = X + 10
        Print "重力:" & Format(GetGravity(Planet), "0.00") & "m/s^2"
        
        CurrentX = X + 10
        Print "建筑:" & PlanetGetUsedBlock(Planet) & "/" & Planet.UtilizingBlock
        
        CurrentX = X + 10
        Print "蒸发量:" & NumberFormat(PlanetGetEvaporation(Planet))
        
        CurrentX = X + 10
        Print "降水密度:" & NumberFormat(PlanetGetRainfallDensity(Planet))
        
        CurrentX = X + 10
        Print "径流量:" & NumberFormat(PlanetGetRunoff(Planet))
        
        DrawButtonWithUI RectangleCreate(X + 240, Y + 40, 80, 40), "增加建筑空间" & vbCrLf & "花费:" & GetBlockCost(Planet), "add_utilizing_block " & SelectPlanet
    Case 1
        DrawButtonWithUI RectangleCreate(X + 10, Y + 40, 60, 20), "基本信息", "system(""set_menu_button 0"")"
        DrawButtonWithUI RectangleCreate(X + 80, Y + 40, 40, 20), "文明", "system(""set_menu_button 1"")", , RGB(128, 128, 128)
        DrawButtonWithUI RectangleCreate(X + 130, Y + 40, 40, 20), "市场", "system(""set_menu_button 2"")"
        
        With Planets(SelectPlanet)
            Print
            CurrentX = X + 10
            Print .Name
            CurrentX = X + 20
            Print "总人口:" & NumberFormat(.Colonys.Population) & "/" & NumberFormat(Planets(SelectPlanet).Housing)
            CurrentX = X + 20
            If Planets(SelectPlanet).Housing < .Colonys.Population Then
                Print "居住空间不足"
            Else
                If Planets(SelectPlanet).Housing = .Colonys.Population Then
                    Print "居住空间已满"
                Else
                    Print "居住空间充足"
                End If
            End If
        
            CurrentX = X + 20
            Print "岗位数:" & GetStaffNeed(Planets(SelectPlanet))
            CurrentX = X + 20
            Print "在岗率:" & Format(GetStaffAdyquacy(Planets(SelectPlanet)), "0.00%")
            
            CurrentX = X + 20
            Print "自然增长:" & NumberFormat(GetGrowthRate(Planets(SelectPlanet)) * .Colonys.Population)
            
            CurrentX = X + 20
            Print GetResourceName(ResourceFood) & "消耗:" & NumberFormat(-.Colonys.Population);
            
            If .Resources(ResourceFood) < .Colonys.Population Then
                Print " 缺乏" & GetResourceName(ResourceFood);
                Print " 因为缺乏" & GetResourceName(ResourceFood) & "而死亡:" & NumberFormat(-.Colonys.Population * 0.02 * (1 - .Resources(ResourceFood) / .Colonys.Population))
            Else
                Print " " & GetResourceName(ResourceFood) & "充足"
            End If
            
            CurrentX = X + 20
            Print GetResourceName(ResourceCleanWater) & "消耗:" & NumberFormat(-.Colonys.Population);
            
            If .Resources(ResourceCleanWater) < .Colonys.Population Then
                Print " 缺乏" & GetResourceName(ResourceCleanWater);
                Print " 因为缺乏" & GetResourceName(ResourceCleanWater) & "而死亡:" & NumberFormat(-.Colonys.Population * 0.02 * (1 - .Resources(ResourceCleanWater) / .Colonys.Population))
            Else
                Print " " & GetResourceName(ResourceCleanWater) & "充足"
            End If
        End With
    Case 2
        DrawButtonWithUI RectangleCreate(X + 10, Y + 40, 60, 20), "基本信息", "system(""set_menu_button 0"")"
        DrawButtonWithUI RectangleCreate(X + 80, Y + 40, 40, 20), "文明", "system(""set_menu_button 1"")"
        DrawButtonWithUI RectangleCreate(X + 130, Y + 40, 40, 20), "市场", "system(""set_menu_button 2"")", , RGB(128, 128, 128)

        For i = 1 To UBound(Planet.Resources)
            FillColor = vbWhite
            Line (X + 10, Y + 20 + 50 * i)-(X + 330, Y + 60 + 50 * i), , B
            
            MyPrint GetResourceName(i), X + 10, Y + 50 * i + 25
            
            CurrentY = Y + 50 * i + 25
            MyPrint "市场存储:" & NumberFormat(Planet.Market.Storage(i)), X + 80, CurrentY
            MyPrint "市场金钱:" & NumberFormat(Planet.Market.Money(i)), X + 80, CurrentY
            If Planet.Market.Prices(i) < 0.01 And Planet.Market.Prices(i) <> 0 Then
                MyPrint "价格:" & Format(Planet.Market.Prices(i), "0.000E-00"), X + 80, CurrentY
            Else
                MyPrint "价格:" & Format(Planet.Market.Prices(i), "0.000"), X + 80, CurrentY
            End If
        Next
    End Select
    
    DrawCloseButton X + 330, Y + 5
End Sub

Private Sub DrawGameUIPopulation(ByVal X As Long, ByVal Y As Long)
    Dim DrawPosition As Rectangle
    Dim i As Long
    
    FillColor = vbWhite
    Line (X, Y)-(X + 400, Y + 300), , B
    
    MyPrint "人口", X + 10, Y + 15
    
    If SelectPlanet = 0 Then
        For i = 1 To UBound(Planets)
            With Planets(i)
                Print
                CurrentX = X + 10
                Print .Name
                CurrentX = X + 20
                Print "总人口:" & NumberFormat(.Colonys.Population) & "/" & NumberFormat(Planets(i).Housing) & " ";
                If Planets(i).Housing < .Colonys.Population Then
                    Print "居住空间不足"
                Else
                    If Planets(i).Housing = .Colonys.Population Then
                        Print "居住空间已满"
                    Else
                        Print "居住空间充足"
                    End If
                End If
            
                CurrentX = X + 20
                Print "岗位数:" & GetStaffNeed(Planets(i))
                CurrentX = X + 20
                Print "在岗率:" & Format(GetStaffAdyquacy(Planets(i)), "0.00%")
                
                CurrentX = X + 20
                Print "自然增长:" & NumberFormat(GetGrowthRate(Planets(i)) * .Colonys.Population)
                
                CurrentX = X + 20
                Print GetResourceName(ResourceFood) & "消耗:" & NumberFormat(-.Colonys.Population);
                
                If .Resources(ResourceFood) < .Colonys.Population Then
                    Print " 缺乏" & GetResourceName(ResourceFood);
                    Print " 因为缺乏" & GetResourceName(ResourceFood) & "而死亡:" & -Int(.Colonys.Population * 0.05)
                Else
                    Print " " & GetResourceName(ResourceFood) & "充足"
                End If
                
                CurrentX = X + 20
                Print GetResourceName(ResourceCleanWater) & "消耗:" & NumberFormat(-.Colonys.Population);
                
                If .Resources(ResourceCleanWater) < .Colonys.Population Then
                    Print " 缺乏" & GetResourceName(ResourceCleanWater);
                    Print " 因为缺乏" & GetResourceName(ResourceCleanWater) & "而死亡:" & -Int(.Colonys.Population * 0.05)
                Else
                    Print " " & GetResourceName(ResourceCleanWater) & "充足"
                End If
            End With
        Next
    Else
        With Planets(SelectPlanet)
            Print
            CurrentX = X + 10
            Print .Name
            CurrentX = X + 20
            Print "总人口:" & NumberFormat(.Colonys.Population) & "/" & NumberFormat(Planets(SelectPlanet).Housing) & " ";
            If Planets(SelectPlanet).Housing < .Colonys.Population Then
                Print "居住空间不足"
            Else
                If Planets(SelectPlanet).Housing = .Colonys.Population Then
                    Print "居住空间已满"
                Else
                    Print "居住空间充足"
                End If
            End If
        
            CurrentX = X + 20
            Print "岗位数:" & GetStaffNeed(Planets(SelectPlanet))
            CurrentX = X + 20
            Print "在岗率:" & Format(GetStaffAdyquacy(Planets(SelectPlanet)), "0.00%")
            
            CurrentX = X + 20
            Print "自然增长:" & NumberFormat(GetGrowthRate(Planets(SelectPlanet)) * .Colonys.Population)
            
            CurrentX = X + 20
            Print GetResourceName(ResourceFood) & "消耗:" & NumberFormat(-.Colonys.Population);
            
            If .Resources(ResourceFood) < .Colonys.Population Then
                Print " 缺乏" & GetResourceName(ResourceFood);
                Print " 因为缺乏" & GetResourceName(ResourceFood) & "而死亡:" & NumberFormat(-.Colonys.Population * 0.05)
            Else
                Print " " & GetResourceName(ResourceFood) & "充足"
            End If
            
            CurrentX = X + 20
            Print GetResourceName(ResourceCleanWater) & "消耗:" & NumberFormat(-.Colonys.Population);
            
            If .Resources(ResourceCleanWater) < .Colonys.Population Then
                Print " 缺乏" & GetResourceName(ResourceCleanWater);
                Print " 因为缺乏" & GetResourceName(ResourceCleanWater) & "而死亡:" & NumberFormat(-.Colonys.Population * 0.05)
            Else
                Print " " & GetResourceName(ResourceCleanWater) & "充足"
            End If
        
            MyPrint "现金:" & NumberFormat(.Colonys.PopulationMoney), X + 20, CurrentY
            MyPrint "信教比例:" & Format(.Colonys.PopulationReligion, "0.00%"), X + 20, CurrentY
            MyPrint "平等度:" & Format(.Colonys.Equality, "0.00%"), X + 20, CurrentY
            
            For i = 1 To UBound(.Colonys.PopulationStorage)
                If .Colonys.PopulationStorage(i) > 0 Then
                    MyPrint "存储" & GetResourceName(i) & ":" & NumberFormat(.Colonys.PopulationStorage(i)), X + 20, CurrentY
                End If
            Next
        End With
        
        DrawButtonWithUI RectangleCreate(X + 160, Y + 40, 60, 40), "发展人口" & vbCrLf & "花费:" & 10, "system(""develop_population " & SelectPlanet & " 10"")", "增加1人口"
    End If
    
    DrawCloseButton X + 330, Y + 5
End Sub


Private Sub DrawUIResearch(ByVal X As Long, ByVal Y As Long)
    Dim DrawPosition As Rectangle
    
    FillColor = vbWhite
    Line (X, Y)-(X + 400, Y + 300), , B
    
    MyPrint "研究", X + 10, Y + 15
    
    DrawCloseButton X + 330, Y + 5
End Sub

Private Sub DrawGameUIResource(Planet As Planet, ByVal X As Long, ByVal Y As Long)
    Dim i As Long
    Dim Count As Long
    Dim DrawPosition As Rectangle
    Dim ButtonSize As Long '物品图标大小
    Dim ButtonNum As Long '每行显示的物品图标数量
    
    FillColor = vbWhite
    Line (X, Y)-(X + 400, Y + 300), , B
    MyPrint "资源", X + 10, Y + 15
    
    Select Case SelectMenuButton
    Case 0
        DrawButtonWithUI RectangleCreate(X + 50, Y + 10, 40, 20), "运输", "system(""set_menu_button 1"")"
        
        ButtonSize = 40
        ButtonNum = 7
        For i = 1 To UBound(Planet.Resources)
            '绘制物品图标
            DrawPosition = RectangleCreate(X + 10 + (ButtonSize + 10) * ((i - 1) Mod ButtonNum), Y + 40 + (ButtonSize + 30) * ((i - 1) \ ButtonNum), ButtonSize, ButtonSize)
            DrawButtonWithUI DrawPosition, GetResourceName(i), , , RGB(128, 192, 128)
            
            '在物品图标下方绘制物品数量
            MyPrint NumberFormat(Planet.Resources(i)), DrawPosition.Left + ButtonSize / 2 - TextWidth(NumberFormat(Planet.Resources(i))) / 2, DrawPosition.Top + ButtonSize
        Next
        MyPrint "资源上限:" & Planet.Storage, X + 10, Y + 280
        
    Case 1
        DrawButtonWithUI RectangleCreate(X + 50, Y + 10, 40, 20), "运输", "system(""set_menu_button 0"")", , RGB(128, 128, 128)

        Count = 0
        For i = 1 To UBound(Planet.Resources)
            Count = Count + 1
            FillColor = vbWhite
            Line (X + 10, Y - 10 + 50 * Count)-(X + 330, Y + 30 + 50 * Count), , B
            
            MyPrint GetResourceName(i), X + 10, Y + 50 * Count - 5
            
            MyPrint "持续运输:" & Planet.Transport(i), X + 80, CurrentY
            
            MyPrint "资金:" & -Planet.Transport(i), X + 80, CurrentY
            
            DrawButtonWithUI RectangleCreate(X + 180, Y - 5 + 50 * Count, 40, 30), "增加", "system(""add_transport_resources " & Planets(SelectPlanet).Tag & " " & i & " 10"")"
            DrawButtonWithUI RectangleCreate(X + 230, Y - 5 + 50 * Count, 40, 30), "减少", "system(""add_transport_resources " & Planets(SelectPlanet).Tag & " " & i & " -10"")"
            DrawButtonWithUI RectangleCreate(X + 280, Y - 5 + 50 * Count, 40, 30), "一次性", "system(""transport_resources " & Planets(SelectPlanet).Tag & " " & i & " 10"")"
        Next
        
'        Count = Count + 1
'        FillColor = vbWhite
'
'        DrawPosition = RectangleCreate(X + 10, Y + 50 * Count, 290, 40)
'        DrawButton DrawPosition, "添加运输线"
'        AddUIObjectList DrawPosition, 1, 8, i
    End Select
    
    DrawCloseButton X + 330, Y + 5
End Sub

'绘制教程界面
Private Sub DrawUITutorial(ByVal X As Long, ByVal Y As Long)
    FillColor = vbWhite
    Line (X, Y)-(X + 400, Y + 300), , B
    
    MyPrint "左上角显示详细信息", X + 10, Y + 15
    MyPrint "点击星球进入星球界面", X + 10, CurrentY
End Sub
    
Private Sub DrawMainMenu() '绘制主菜单到屏幕
    Dim ButtonPosition As Rectangle

    Font = "微软雅黑"
    ForeColor = vbYellow
    FontSize = 28
    MyPrint "宜居", GetCenterX, GetCenterY - 110, 1
    
    ForeColor = vbBlack
    FillStyle = 0
    FontSize = 9
    Font = "宋体"
    ButtonPosition = RectangleCreate(GetCenterX - 90, GetCenterY - 40, 180, 40)
    DrawButtonWithUI ButtonPosition, "开始游戏", "system(""change_window " & FormStartGame & """)"
    
    ButtonPosition = RectangleCreate(GetCenterX - 90, GetCenterY + 20, 180, 40)
    DrawButtonWithUI ButtonPosition, "设置", "system(""change_window " & FormSettings & """)"
    
    ButtonPosition = RectangleCreate(GetCenterX - 90, GetCenterY + 80, 180, 40)
    DrawButtonWithUI ButtonPosition, "退出", "system(""change_window " & FormClosed & """)"
    
    MyPrint "宜居 Demo20220722a", 0, ScaleHeight - TextHeight("宜居")
End Sub

Private Sub DrawMenuUI() '绘制菜单UI
    If ShowingUI <> ShowUINone Then
        Select Case ShowingUI
        Case ShowUILoadGame
            DrawUILoadGame 340, 20
        Case ShowUINewGame
            DrawUINewGame 340, 20
        Case ShowUINewCampaign
            DrawUINewCampaign 340, 20
        Case ShowUITutorial
            DrawUITutorial 340, 20
        End Select
    End If
End Sub

Private Sub DrawModules(Planet As Planet)
    Dim i As Long
    Dim MoudlePosition As Rectangle
    Dim MoudleText As String
    Dim DrawPosition As Rectangle
    
    MoudlePosition = RectangleCreate(150, 10, 100, 30)
    
    With MoudlePosition
        '翻页键
        DrawButtonWithUI RectangleCreate(.Left + .Width + 10, .Top, .Height, .Height), "上翻", "system(""change_module_offset " & -300 & """)"
        DrawButtonWithUI RectangleCreate(.Left + .Width + 10, .Top + .Height + 10, .Height, .Height), "下翻", "system(""change_module_offset " & 300 & """)"
        
        DrawPosition = MoudlePosition
        DrawPosition.Top = DrawPosition.Top - DrawModuleOffset
        
        For i = 1 To UBound(Planet.Modules)
            '显示模块选中后的黄色提示框
            If i = SelectModule Then
                RectangleDraw RectangleCreate(DrawPosition.Left - 3, DrawPosition.Top - 3, DrawPosition.Width + 6, DrawPosition.Height + 6), vbYellow
            End If
            
            '显示模块
            If Planet.Modules(i).Construction < 1 Then
                MoudleText = ModuleTypes(Planet.Modules(i).Type).Name & "(建造中)"
            Else
                If Planet.Modules(i).Enabled = True Then
                    MoudleText = ModuleTypes(Planet.Modules(i).Type).Name
                Else
                    MoudleText = ModuleTypes(Planet.Modules(i).Type).Name & "(已禁用)"
                End If
            End If
            DrawButtonWithUI DrawPosition, MoudleText, "system(""switch_select_module " & i & """)"
            DrawPosition.Top = DrawPosition.Top + DrawPosition.Height + 10
        Next
        
        '绘制建造模块按钮
        For i = 1 To UBound(ModuleTypes)
            DrawButtonWithUI DrawPosition, "建造" & ModuleTypes(i).Name, "system(""bulid_module " & Planets(SelectPlanet).Tag & " " & i & """)"
            DrawPosition.Top = DrawPosition.Top + DrawPosition.Height + 10
        Next
    End With
End Sub

Private Sub DrawPauseMenu() '绘制暂停菜单
    Dim ButtonPosition As Rectangle
    
    FillColor = vbWhite
    Line (GetCenterX - 150, GetCenterY - 100)-(GetCenterX + 150, GetCenterY + 100), , B
    
    DrawButtonWithUI RectangleCreate(GetCenterX - 90, GetCenterY - 80, 180, 40), "继续游戏", "system(""swith_pause_menu" & """)"
    
    DrawButtonWithUI RectangleCreate(GetCenterX - 90, GetCenterY - 20, 180, 40), "保存游戏", "system(""save_game" & """)"
    
'    ButtonPosition = RectangleCreate(GetCenterX - 90, GetCenterY - 20, 180, 40)
'    DrawButton ButtonPosition, "设置"
'    AddUIObjectList ButtonPosition, 1, 10, FormSettings

    DrawButtonWithUI RectangleCreate(GetCenterX - 90, GetCenterY + 40, 180, 40), "返回主菜单", "change_window " & FormMainMenu
End Sub

Private Sub DrawProgressBar(Position As Rectangle, ByVal Percent As Double, Optional ByVal Color As OLE_COLOR = &HC0C0C0)
    Dim Text As String
    
    FillColor = vbWhite
    Line (Position.Left, Position.Top)-(Position.Left + Position.Width - 1, Position.Top + Position.Height - 1), , B
    FillColor = Color
    Line (Position.Left + 1, Position.Top + 1)-(Position.Left + 1 + (Position.Width - 3) * Percent, Position.Top + Position.Height - 2), , B
    FillStyle = 1
    Line (Position.Left, Position.Top)-(Position.Left + Position.Width - 1, Position.Top + Position.Height - 1), , B
    FillStyle = 0
    
    Text = Format(Percent, "0.00%")
    MyPrint Text, Position.Left + 0.5 * Position.Width, Position.Top + 0.5 * Position.Height, 1
End Sub

Private Sub DrawSettings() '绘制设置界面
    Dim Position As Rectangle
    
    ReDim UIObjectList(0)
    CurrentX = 80
    CurrentY = 40
    Print "建设行星"
    Position = RectangleCreate(GetCenterX - 180, ScaleHeight - 80, 360, 50)
    DrawButton Position, "完成"
    AddUIObjectList Position, "system(""change_window " & FormMainMenu & """)"
End Sub

'绘制航天器界面
Private Sub DrawSpacecrafts(Planet As Planet)
    Dim i As Long
    Dim MoudlePosition As Rectangle
    Dim DrawPosition As Rectangle
    
    MoudlePosition = RectangleCreate(300, 10, 100, 30)
    
    With MoudlePosition
        DrawPosition = MoudlePosition
        
        For i = 1 To UBound(Spacecrafts)
            If SpacecraftGetPlanet(Spacecrafts(i)) = PlanetGetID(Planet) Then
                '显示飞船选中后的黄色提示框
                If i = SelectSpacecraft Then
                    RectangleDraw RectangleCreate(DrawPosition.Left - 3, DrawPosition.Top - 3, DrawPosition.Width + 6, DrawPosition.Height + 6), vbYellow
                End If
                
                DrawButtonWithUI DrawPosition, Spacecrafts(i).Name, "system(""switch_select_spacecraft " & i & """)"
                DrawPosition.Top = DrawPosition.Top + DrawPosition.Height + 10
            End If
        Next
    End With
End Sub

'绘制游戏速度控制界面
Private Sub DrawSpeedControlBar(Position As Rectangle)
    ForeColor = vbBlack
    FillStyle = 0
    
    With Position
        If GameSpeed <= 0 Then
            DrawButtonWithUI Position, "暂停", "", , RGB(128, 128, 128)
        Else
            DrawButtonWithUI Position, "暂停", "system(""set_speed 0"")", , RGB(192, 192, 192)
        End If
        .Left = .Left + .Width + 10
        
        If GameSpeed = 1 Then
            DrawButtonWithUI Position, "1×", "", , RGB(128, 128, 128)
        Else
            DrawButtonWithUI Position, "1×", "system(""set_speed 1"")", , RGB(192, 192, 192)
        End If
        .Left = .Left + .Width + 10
        
        If GameSpeed = 2 Then
            DrawButtonWithUI Position, "2×", "", , RGB(128, 128, 128)
        Else
            DrawButtonWithUI Position, "2×", "system(""set_speed 2"")", , RGB(192, 192, 192)
        End If
        .Left = .Left + .Width + 10
        
        If GameSpeed = 4 Then
            DrawButtonWithUI Position, "4×", "", , RGB(128, 128, 128)
        Else
            DrawButtonWithUI Position, "4×", "system(""set_speed 4"")", , RGB(192, 192, 192)
        End If
    End With
End Sub

Private Sub DrawStartGame() '绘制开始游戏菜单
    Dim ButtonPosition As Rectangle

    Font = "微软雅黑"
    ForeColor = vbBlack
    FillStyle = 0
    FontSize = 9
    Font = "宋体"
    ButtonPosition = RectangleCreate(40, 40, 180, 40)
    DrawButton ButtonPosition, "载入游戏"
'    AddUIObjectList "Button", ButtonPosition, 1, 3, ShowUILoadGame
    
    DrawButtonWithUI RectangleCreate(40, 100, 180, 40), "新的游戏", "system(""change_window " & FormGame & """)"
    
    ButtonPosition = RectangleCreate(40, 160, 180, 40)
    DrawButton ButtonPosition, "场景"
'    AddUIObjectList ButtonPosition, 1, 3, ShowUINewCampaign
    
    ButtonPosition = RectangleCreate(40, 220, 180, 40)
    DrawButton ButtonPosition, "教程"
'    AddUIObjectList ButtonPosition, 1, 3, ShowUITutorial
    
    DrawMenuUI
End Sub

Private Sub DrawToolTip(ByVal X As Long, ByVal Y As Long, ByVal Text As String) '绘制提示信息
    ForeColor = vbBlack
    FillStyle = 0
    FillColor = vbWhite
    Line (X, Y)-(X + TextWidth(Text) + 10, Y + TextHeight(Text) + 10), , B
    MyPrint Text, X + 5, Y + 5
End Sub

Private Sub DrawToolTips() '检查并绘制当前所有提示信息
    Dim MouseX As Long, MouseY As Long '当前鼠标的x和y坐标
    Dim i As Long
    
    MouseX = MousePosition.X
    MouseY = MousePosition.Y
    For i = UBound(UIObjectList) To 1 Step -1
        If MouseX >= UIObjectList(i).Position.Left And MouseX <= UIObjectList(i).Position.Left + UIObjectList(i).Position.Width And MouseY >= UIObjectList(i).Position.Top And MouseY <= UIObjectList(i).Position.Top + UIObjectList(i).Position.Height Then
            If UIObjectList(i).Tooltip <> "" Then
                DrawToolTip MouseX, MouseY, GetTooltip(UIObjectList(i))
            End If
            Exit For
        End If
    Next i
End Sub

'效果计算
Private Sub EffectCalculate(Effect As Effect, Planet As Planet, Efficiency As Double)
    With Effect
        Select Case .Type
        Case EffectOxygen
            PlanetChangeMaterial Planet, "氧气", .Amont * Efficiency
        Case EffectGas
            PlanetChangeMaterial Planet, "稳定气体", .Amont * Efficiency
        Case EffectWater
            PlanetChangeMaterial Planet, "水", .Amont * Efficiency
        Case EffectTempreture
            Planet.Tempreture = Planet.Tempreture + .Amont * Efficiency
        Case EffectResource
            If Not Planet.HomeWorld Then
                PlanetAddResource Planet, .EffectResources.Type, .EffectResources.Amont * Efficiency
            Else
                PlanetAddResource Planet, .EffectResources.Type, .EffectResources.Amont * Efficiency
                Planet.Market.Storage(.EffectResources.Type) = Planet.Market.Storage(.EffectResources.Type) + .EffectResources.Amont * Efficiency
            End If
        Case EffectResearchPoint
            ResearchPoint = ResearchPoint + .Amont * Efficiency
        Case EffectPrestige
            Prestige = Prestige + .Amont * Efficiency
        Case EffectHousing
            Planet.Housing = Planet.Housing + .Amont * Efficiency
        Case EffectStorage
            Planet.Storage = Planet.Storage + .Amont * Efficiency
        End Select
    End With
End Sub

Private Function EffectCreate(EffectType As EffectEnum, EffectResources As Resource, Optional ByVal Amont As Double = 0) As Effect
    With EffectCreate
        .Type = EffectType
        .Amont = Amont
        .EffectResources = EffectResources
    End With
End Function

Private Function EffectNull() As Effect
    EffectNull = EffectCreate(EffectNone, ResourceNull)
End Function

Private Sub ExpansionModule(Planet As Planet, ByVal n As Long) '扩建建筑
    Dim i As Long
    
    If Planet.HomeWorld Then
        MsgBox "建筑无法建设在母星"
        Exit Sub
    End If
    
    With ModuleTypes(Planet.Modules(n).Type)
        '检查资金是否足够
        If Money < .Cost Then
            MsgBox "缺少" & NumberFormat(.Cost - Money) & "资金"
            Exit Sub
        End If
        
        '检查空间是否足够
        If .Space > 0 Then
            If PlanetGetUsedBlock(Planet) + .Space > Planet.UtilizingBlock Then
                MsgBox Planet.Name & "缺少" & NumberFormat(PlanetGetUsedBlock(Planet) + .Space - Planet.UtilizingBlock) & "建筑空间"
                Exit Sub
            End If
        End If
        
        '检查资源是否足够
        For i = 1 To UBound(.Resources)
            If Planet.Resources(.Resources(i).Type) < .Resources(i).Amont Then
                MsgBox "缺少" & NumberFormat(.Resources(i).Amont - Planet.Resources(.Resources(i).Type)) & GetResourceName(.Resources(i).Type)
                Exit Sub
            End If
        Next
        
        Planet.Modules(n).Construction = 0
    
        '扣除资源
        Money = Money - .Cost
        For i = 1 To UBound(.Resources)
            PlanetAddResource Planet, .Resources(i).Type, -.Resources(i).Amont
        Next
    End With
End Sub

Private Function FinancialInfoCreate() As FinancialInfo
    With FinancialInfoCreate
        .Funds = 0
        .Contribution = 0
        .Income = 0
        .ColonyMaintence = 0
        .Salary = 0
        .Transport = 0
        .Expence = 0
        .NetIncome = 0
    End With
End Function

Private Function FindPlanetWithTag(Tag As String) As Long
    Dim i As Long
    For i = 1 To UBound(Planets)
        If Planets(i).Tag = Tag Then
            FindPlanetWithTag = i
        End If
    Next
End Function

Private Function FindStarWithTag(Tag As String) As Long
    Dim i As Long
    For i = 1 To UBound(Stars)
        If Stars(i).Tag = Tag Then
            FindStarWithTag = i
        End If
    Next
End Function

Private Sub GameInitialization()
    SelectModule = 0
    IsShowPauseMenu = False
    LoadMaterial
    LoadModule
    LoadEvent
    LoadSystem System
    LoadSpaceCraft
    GameDate = #1/1/2200#
    Money = 1000
    Funds = 100
    FinancialInfo = FinancialInfoCreate
    PreviousFinancialInfo = FinancialInfoCreate
    ReDim ShowEvents(1)
    ShowEvents(1) = 1
End Sub

Private Function GetCenterX() As Long
    GetCenterX = ScaleWidth / 2
End Function

Private Function GetCenterY() As Long
    GetCenterY = ScaleHeight / 2
End Function

Private Function GetDpi()
    GetDpi = Screen.TwipsPerPixelX
End Function

Private Function GetEffectName(Effect As EffectEnum)
    Select Case Effect
    Case EffectNone
        GetEffectName = "无效果"
    Case EffectOxygen
        GetEffectName = "氧气"
    Case EffectGas
        GetEffectName = "气压"
    Case EffectWater
        GetEffectName = "水"
    Case EffectTempreture
        GetEffectName = "温度"
    Case EffectReflectivity
        GetEffectName = "反射率"
    Case EffectHousing
        GetEffectName = "住房"
    Case EffectResource
        GetEffectName = "资源"
    Case EffectResearchPoint
        GetEffectName = "研究点数"
    Case EffectPrestige
        GetEffectName = "声望"
    Case EffectStorage
        GetEffectName = "存储空间"
    Case EffectSolarPower
        GetEffectName = "太阳能"
    Case EffectRunOff
        GetEffectName = "径流量"
    End Select
End Function

Private Function GetGravity(Planet As Planet) As Double
    GetGravity = PlanetGetMass(Planet) / Planet.Radio ^ 2 * 6.67259 * 10 ^ -11 * 10 ^ 13
End Function

Private Sub GetFPS()
    Static NowTime As String
    Static FpsNew As Long
    If NowTime <> Str(Time) Then
        NowTime = Str(Time)
        Fps = FpsNew
        FpsNew = 1
    Else
        FpsNew = FpsNew + 1
    End If
End Sub

Private Function GetFrameWidth()
    GetFrameWidth = (Width / GetDpi - ScaleWidth) / 2
End Function

Private Function GetFrameTop()
    GetFrameTop = Height / GetDpi - ScaleHeight - GetFrameWidth
End Function

Private Function GetBlockCost(Planet As Planet) As Double
    GetBlockCost = Int(Log((Planet.UtilizingBlock + 1) / PlanetGetBlock(Planet) + 1) * 1000)
End Function

Private Function GetGrowthRate(Planet As Planet) As Double '计算行星的自然增长率
    With Planet
        If Planet.Housing < .Colonys.Population Then
            GetGrowthRate = -0.01 * (.Colonys.Population - Planet.Housing) / .Colonys.Population
        Else
            If Planet.Housing = .Colonys.Population Then
                GetGrowthRate = 0
            Else
                GetGrowthRate = 0.001
            End If
        End If
        
        If .Resources(ResourceFood) < .Colonys.Population * 0.0001 Then
            GetGrowthRate = GetGrowthRate - 0.02 * (1 - .Resources(ResourceFood) / (.Colonys.Population * 0.0001))
        End If
        
        If .Resources(ResourceCleanWater) < .Colonys.Population * 0.0001 Then
            GetGrowthRate = GetGrowthRate - 0.02 * (1 - .Resources(ResourceCleanWater) / (.Colonys.Population * 0.0001))
        End If
    End With
End Function

Private Function GetPowerAdyquacy(Planet As Planet) As Double
    If GetPowerProduce(Planet) >= GetPowerUse(Planet) Then
        GetPowerAdyquacy = 1
    Else
        GetPowerAdyquacy = GetPowerProduce(Planet) / GetPowerUse(Planet)
    End If
End Function

Private Function GetPowerProduce(Planet As Planet) As Double
    Dim i As Long

    GetPowerProduce = 0
    For i = 1 To UBound(Planet.Modules)
        With Planet.Modules(i)
            If ModuleTypes(.Type).Power > 0 Then
                GetPowerProduce = GetPowerProduce + ModuleTypes(.Type).Power * .Efficiency * .Size
            End If
        End With
    Next

    For i = 1 To UBound(Spacecrafts)
        With Spacecrafts(i)
            If SpacecraftGetPlanet(Spacecrafts(i)) = PlanetGetID(Planet) Then GetPowerProduce = GetPowerProduce + .Power
        End With
    Next
End Function

Private Function GetPowerUse(Planet As Planet) As Double
    Dim i As Long
    GetPowerUse = 0

    For i = 1 To UBound(Planet.Modules)
        With Planet.Modules(i)
            If ModuleTypes(.Type).Power < 0 Then
                GetPowerUse = GetPowerUse - ModuleTypes(.Type).Power * .EfficiencyModifier * .Size
            End If
        End With
    Next
End Function

Private Function GetReflectivity(Planet As Planet) As Double
    Dim i As Long
    Dim j As Long
    
    GetReflectivity = Planet.Reflectivity
    For i = 1 To UBound(Planet.Modules)
        With Planet.Modules(i)
            For j = 1 To UBound(ModuleTypes(.Type).Effects)
                If ModuleTypes(.Type).Effects(j).Type = EffectReflectivity Then
                    GetReflectivity = GetReflectivity + (ModuleTypes(.Type).Effects(j).Amont - Planet.Reflectivity) * .Efficiency / (UBound(Planet.Modules) + 10)
                End If
            Next
        End With
    Next
End Function

Private Function GetResourceName(ResourceEnum As ResourceEnum)
    Select Case ResourceEnum
    Case ResourceNone
        GetResourceName = "无资源"
    Case ResourceRock
        GetResourceName = "岩石"
    Case ResourceMineral
        GetResourceName = "矿物"
    Case ResourceMetel
        GetResourceName = "金属"
    Case ResourceComposites
        GetResourceName = "复合材料"
    Case ResourceFood
        GetResourceName = "食物"
    Case ResourceCleanWater
        GetResourceName = "净水"
    Case ResourceRobot
        GetResourceName = "机器人"
    Case ResourceFurniture
        GetResourceName = "家具"
    Case ResourceElectricAppliance
        GetResourceName = "家电"
    Case ResourceBioMaterial
        GetResourceName = "生物材料"
    End Select
End Function

Private Function GetStaffAdyquacy(Planet As Planet) As Double
    If Planet.Colonys.Population >= GetStaffNeed(Planet) Then
        GetStaffAdyquacy = 1
    Else
        GetStaffAdyquacy = Planet.Colonys.Population / GetStaffNeed(Planet)
    End If
End Function

Private Function GetStaffNeed(Planet As Planet) As Double
    Dim i As Long
    GetStaffNeed = 0
    For i = 1 To UBound(Planet.Modules)
        With Planet.Modules(i)
            If ModuleTypes(.Type).Staff > 0 Then
                GetStaffNeed = GetStaffNeed + ModuleTypes(.Type).Staff * .EfficiencyModifier * .Size
            End If
        End With
    Next
End Function

Private Function GetTooltip(UIObject As UIObject) As String '获得提示信息
    Dim i As Long
    
    GetTooltip = ""
    Select Case UIObject.ClickAction
    Case 0 '无效果
    Case 1 '发展人口
        GetTooltip = "增加1人口"
    Case 2 '建造建筑
'        With ModuleTypes(UIObject.ClickAddedCode)
'            GetTooltip = "建造" & .Name
'            GetTooltip = GetTooltip & vbCrLf & .Description
'
'            If Money >= .Cost Then
'                GetTooltip = GetTooltip & vbCrLf & "花费:" & .Cost
'            Else
'                GetTooltip = GetTooltip & vbCrLf & "花费:" & .Cost & "(不足)"
'            End If
'            GetTooltip = GetTooltip & vbCrLf & "建造时间:" & .BuildTime & "天"
'            If PlanetGetUsedBlock(Planets(SelectPlanet)) + .Space <= Planets(SelectPlanet).UtilizingBlock Then
'                GetTooltip = GetTooltip & vbCrLf & "建筑空间:" & .Space
'            Else
'                GetTooltip = GetTooltip & vbCrLf & "建筑空间:" & .Space & "(不足)"
'            End If
'
'            For i = 1 To UBound(.Resources)
'                GetTooltip = GetTooltip & vbCrLf & ResourceToString(.Resources(i))
'                With .Resources(i)
'                    If Planets(SelectPlanet).Resources(.Type) < .Amont Then
'                        GetTooltip = GetTooltip & "(不足)"
'                    End If
'                End With
'            Next
'        End With
        
    Case 3 '展示UI
    Case 4 '显示建筑
'        With Planets(SelectPlanet).Modules(UIObject.ClickAddedCode)
'            GetTooltip = ModuleTypes(.Type).Name & " 规模:" & .Size
'            GetTooltip = GetTooltip & vbCrLf & ModuleTypes(.Type).Description
'        End With
    Case 5 '建筑翻页
    Case 6 '拆除建筑
'        GetTooltip = "拆除" & ModuleTypes(Planets(SelectPlanet).Modules(UIObject.ClickAddedCode).Type).Name
    Case 7 '一次性运输资源
    Case 8 '增加持续运输资源
    Case 9 '减少持续运输资源
    Case 10 '切换窗口
    Case 11 '增加建筑空间
        GetTooltip = "增加1建筑空间"
    Case 12 '选择星球
'        If UIObject.ClickAddedCode = 0 Then
'            GetTooltip = "显示星系"
'        Else
'            GetTooltip = Planets(UIObject.ClickAddedCode).Name
'        End If
        
    Case 13 '关闭事件
    Case 14 '改变游戏速度
    Case 15 '显示/隐藏暂停界面
    Case 16 '启/禁用建筑
    Case 17 '扩建建筑
'        With ModuleTypes(Planets(SelectPlanet).Modules(UIObject.ClickAddedCode).Type)
'            GetTooltip = "扩建" & .Name
'
'            If Money >= .Cost Then
'                GetTooltip = GetTooltip & vbCrLf & "花费:" & .Cost
'            Else
'                GetTooltip = GetTooltip & vbCrLf & "花费:" & .Cost & "(不足)"
'            End If
'
'            If PlanetGetUsedBlock(Planets(SelectPlanet)) + .Space <= Planets(SelectPlanet).UtilizingBlock Then
'                GetTooltip = GetTooltip & vbCrLf & "建筑空间:" & .Space
'            Else
'                GetTooltip = GetTooltip & vbCrLf & "建筑空间:" & .Space & "(不足)"
'            End If
'
'            For i = 1 To UBound(.Resources)
'                GetTooltip = GetTooltip & vbCrLf & ResourceToString(.Resources(i))
'                With .Resources(i)
'                    If Planets(SelectPlanet).Resources(.Type) < .Amont Then
'                        GetTooltip = GetTooltip & "(不足)"
'                    End If
'                End With
'            Next
'        End With
    Case 18 '改变界面按钮
    Case 19 '点击仓库资源
'        GetTooltip = GetResourceName(UIObject.ClickAddedCode) & ":" & NumberFormat(Planets(SelectPlanet).Resources(UIObject.ClickAddedCode)) & "/" & NumberFormat(GetStorage(Planets(SelectPlanet)))
'        GetTooltip = GetTooltip & vbCrLf & "损耗:" & NumberFormat(Min(-0.01 * (Planets(SelectPlanet).Resources(UIObject.ClickAddedCode) - GetStorage(Planets(SelectPlanet))), 0))
    End Select
End Function

Private Sub LoadEvent()
    ReDim Events(1)
    
    With Events(1)
        .Title = "你好"
        .Content = "欢迎进入游戏。按空格键暂停游戏。按ESC键显示/隐藏" & vbCrLf & "暂停菜单"
        ReDim .Options(1)
        .Options(1) = "确定"
    End With
End Sub

Private Sub LoadMaterial()
    ReDim MaterialTypes(25)
    
    With MaterialTypes(1)
        .Name = "冰"
        .HighTemp = 273.15
        .HighTempTarget = "水"
        .State = MaterialStateSolid
    End With
    
    With MaterialTypes(2)
        .Name = "水"
        .LowTemp = 273.15
        .LowTempTarget = "冰"
        .HighTemp = 373.15
        .HighTempTarget = "水蒸气"
        .State = MaterialStateLiquid
    End With
    
    With MaterialTypes(3)
        .Name = "水蒸气"
        .LowTemp = 373.15
        .LowTempTarget = "水"
        .State = MaterialStateGas
    End With
    
    With MaterialTypes(4)
        .Name = "固态氧"
        .HighTemp = 54.3
        .HighTempTarget = "液态氧"
        .State = MaterialStateSolid
    End With
    
    With MaterialTypes(5)
        .Name = "液态氧"
        .LowTemp = 54.3
        .LowTempTarget = "固态氧"
        .HighTemp = 90.15
        .HighTempTarget = "氧气"
        .State = MaterialStateLiquid
    End With
    
    With MaterialTypes(6)
        .Name = "氧气"
        .LowTemp = 90.15
        .LowTempTarget = "液态氧"
        .State = MaterialStateGas
    End With
    
    With MaterialTypes(7)
        .Name = "固态氢"
        .HighTemp = 14
        .HighTempTarget = "液态氢"
        .State = MaterialStateSolid
    End With
    
    With MaterialTypes(8)
        .Name = "液态氢"
        .LowTemp = 14
        .LowTempTarget = "固态氢"
        .HighTemp = 21
        .HighTempTarget = "氢气"
        .State = MaterialStateLiquid
    End With
    
    With MaterialTypes(9)
        .Name = "氢气"
        .LowTemp = 21
        .LowTempTarget = "液态氢"
        .State = MaterialStateGas
    End With
    
    With MaterialTypes(10)
        .Name = "固态稳定气体"
        .HighTemp = 62
        .HighTempTarget = "液态稳定气体"
        .State = MaterialStateSolid
    End With
    
    With MaterialTypes(11)
        .Name = "液态稳定气体"
        .LowTemp = 62
        .LowTempTarget = "固态稳定气体"
        .HighTemp = 77
        .HighTempTarget = "稳定气体"
        .State = MaterialStateLiquid
    End With
    
    With MaterialTypes(12)
        .Name = "稳定气体"
        .LowTemp = 77
        .LowTempTarget = "液态稳定气体"
        .State = MaterialStateGas
    End With
    
    With MaterialTypes(13)
        .Name = "固态碳氧化物"
        .HighTemp = 195
        .HighTempTarget = "碳氧化物气体"
        .State = MaterialStateSolid
    End With
    
    With MaterialTypes(15)
        .Name = "碳氧化物气体"
        .LowTemp = 195
        .LowTempTarget = "固态碳氧化物"
        .State = MaterialStateGas
    End With
    
    With MaterialTypes(16)
        .Name = "固态烃类"
        .HighTemp = 85
        .HighTempTarget = "液态烃类"
        .State = MaterialStateSolid
    End With
    
    With MaterialTypes(17)
        .Name = "液态烃类"
        .LowTemp = 85
        .LowTempTarget = "固态烃类"
        .HighTemp = 231
        .HighTempTarget = "气态烃类"
        .State = MaterialStateLiquid
    End With
    
    With MaterialTypes(18)
        .Name = "气态烃类"
        .LowTemp = 231
        .LowTempTarget = "液态烃类"
        .State = MaterialStateGas
    End With
    
    With MaterialTypes(19)
        .Name = "岩石"
        .HighTemp = 1670
        .HighTempTarget = "岩浆"
        .State = MaterialStateSolid
    End With
    
    With MaterialTypes(20)
        .Name = "岩浆"
        .LowTemp = 1670
        .LowTempTarget = "岩石"
        .HighTemp = 2630
        .HighTempTarget = "气态岩"
        .State = MaterialStateLiquid
    End With
    
    With MaterialTypes(21)
        .Name = "气态岩"
        .LowTemp = 2630
        .LowTempTarget = "岩浆"
        .State = MaterialStateGas
    End With
    
    With MaterialTypes(22)
        .Name = "金属"
        .HighTemp = 1808
        .HighTempTarget = "液态金属"
        .State = MaterialStateSolid
    End With
    
    With MaterialTypes(23)
        .Name = "液态金属"
        .LowTemp = 1808
        .LowTempTarget = "金属"
        .HighTemp = 3023
        .HighTempTarget = "气态金属"
        .State = MaterialStateLiquid
    End With
    
    With MaterialTypes(24)
        .Name = "气态金属"
        .LowTemp = 3023
        .LowTempTarget = "液态金属"
        .State = MaterialStateGas
    End With
    
    With MaterialTypes(25)
        .Name = "泥土"
        .HighTemp = 1670
        .HighTempTarget = "岩浆"
        .State = MaterialStateSolid
    End With
End Sub

Private Sub LoadModule()
    Dim Effects() As Effect
    Dim Resources() As Resource
    Dim i As Long
    
    ReDim ModuleTypes(0)
'    '农田
'    ModuleTypes(5) = MoudleTypeCreate("农田", "产出食物。", 100, 1, False, -10, 10, 0, 60)
'    With ModuleTypes(5)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceComposites, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectResource
'        .Effects(1).EffectResources = ResourceCreate(ResourceFood, 30)
'    End With
    
    '农田(小)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 0.05))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("农田(小)", Effects, Resources, "小型农田，产出食物。", 100, 0.1, False, -1, 100, -0.02, 60)
    
    '农田(中)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 55))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("农田(中)", Effects, Resources, "中型农田，产出食物。", 100000, 100, False, -1000, 100000, -20, 120)
    
    '农田(大)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 60000))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("农田(大)", Effects, Resources, "大型农田，产出食物。", 100000000, 100000, False, -1000000, 100000000, -20000, 180)

'    '传统农业
'    ModuleTypes(21) = MoudleTypeCreate("传统农业", "使用耕地。产出食物", 100, 8, True, -30, 50000000, -10000000, 300)
'    With ModuleTypes(21)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceComposites, 20000)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectResource
'        .Effects(1).EffectResources = ResourceCreate(ResourceFood, 100000000)
'    End With
    
    '传统农业(小)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(0)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 0.05))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("传统农业(小)", Effects, Resources, "小型传统农业，使用大量人力。产出食物。", 0.1, 1, True, 0, 100, 0, 60)
    
    '传统农业(中)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(0)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 55))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("传统农业(中)", Effects, Resources, "中型传统农业，使用大量人力。产出食物。", 100, 1000, True, 0, 100000, 0, 120)
    
    '传统农业(大)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(0)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 60000))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("传统农业(大)", Effects, Resources, "大型传统农业，使用大量人力。产出食物。", 100000, 1000000, True, 0, 100000000, 0, 180)
    
    '现代农业(小)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 0.05))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("现代农业(小)", Effects, Resources, "小型现代农业。使用耕地，产出食物。", 10, 10, True, -0.8, 100, -0.1, 100)
    
    '现代农业(中)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 55))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("现代农业(中)", Effects, Resources, "中型现代农业。使用耕地，产出食物。", 10000, 10000, True, -800, 100000, -100, 200)
    
    '现代农业(大)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 60000))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("现代农业(大)", Effects, Resources, "大型现代农业。使用耕地，产出食物。", 10000000, 10000000, True, -800000, 100000000, -100000, 300)
    
'    '科考站
'    ModuleTypes(1) = MoudleTypeCreate("科考站", "自给自足的小型站点。可以产出少量科研点数。", 500, 10, False, -50, 0, 0, 60)
'    With ModuleTypes(1)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceComposites, 10)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(5)
'        .Effects(1).Type = EffectHousing
'        .Effects(1).Amont = 15
'        .Effects(2).Type = EffectResearchPoint
'        .Effects(2).Amont = 0.1
'        .Effects(3).Type = EffectStorage
'        .Effects(3).Amont = 50
'        .Effects(4).Type = EffectResource
'        .Effects(4).EffectResources = ResourceCreate(ResourceFood, 15)
'        .Effects(5).Type = EffectResource
'        .Effects(5).EffectResources = ResourceCreate(ResourceCleanWater, 15)
'    End With

    '科考站(小)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(2)
    Resources(1) = ResourceCreate(ResourceComposites, 10)
    Resources(2) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(5)
    Effects(1) = EffectCreate(EffectHousing, ResourceNull, 150)
    Effects(2) = EffectCreate(EffectResearchPoint, ResourceNull, 0.1)
    Effects(3) = EffectCreate(EffectStorage, ResourceNull, 0.05)
    Effects(4) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 0.015))
    Effects(5) = EffectCreate(EffectResource, ResourceCreate(ResourceCleanWater, 0.015))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("科考站(小)", Effects, Resources, "自给自足的小型站点。可以产出少量科研点数。", 10, 10, False, -0.8, 100, -0.1, 100)
    
    '科考站(中)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(2)
    Resources(1) = ResourceCreate(ResourceComposites, 10000)
    Resources(2) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(5)
    Effects(1) = EffectCreate(EffectHousing, ResourceNull, 150000)
    Effects(2) = EffectCreate(EffectResearchPoint, ResourceNull, 100)
    Effects(3) = EffectCreate(EffectStorage, ResourceNull, 50)
    Effects(4) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 15))
    Effects(5) = EffectCreate(EffectResource, ResourceCreate(ResourceCleanWater, 15))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("科考站(中)", Effects, Resources, "自给自足的中型站点。可以产出少量科研点数。", 10000, 10000, False, -800, 100000, -100, 200)
    
    '科考站(大)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(2)
    Resources(1) = ResourceCreate(ResourceComposites, 10000000)
    Resources(2) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(5)
    Effects(1) = EffectCreate(EffectHousing, ResourceNull, 150000000)
    Effects(2) = EffectCreate(EffectResearchPoint, ResourceNull, 100000)
    Effects(3) = EffectCreate(EffectStorage, ResourceNull, 50000)
    Effects(4) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 15000))
    Effects(5) = EffectCreate(EffectResource, ResourceCreate(ResourceCleanWater, 15000))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("科考站(大)", Effects, Resources, "自给自足的大型站点。可以产出少量科研点数。", 10000000, 10000000, False, -800000, 100000000, -100000, 300)
    
'    '太阳能电池板
'    ModuleTypes(3) = MoudleTypeCreate("太阳能电池板", "提供电力，需要太阳光。", 100, 1, False, -10, 0, 100, 60)
'    With ModuleTypes(3)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceComposites, 10)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectSolarPower
'    End With

    '太阳能发电站(小)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectSolarPower, ResourceNull)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("太阳能发电站(小)", Effects, Resources, "小型太阳能发电站。提供电力，需要光照。", 100, 10, False, -1, 100, 0.1, 60)
    
    '太阳能发电站(中)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectSolarPower, ResourceNull)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("太阳能发电站(中)", Effects, Resources, "中型太阳能发电站。提供电力，需要光照。", 100000, 10000, False, -1000, 100000, 100, 120)
    
    '太阳能发电站(大)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectSolarPower, ResourceNull)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("太阳能发电站(大)", Effects, Resources, "大型太阳能发电站。提供电力，需要光照。", 100000000, 10000000, False, -1000000, 100000000, 100000, 180)

'    '火力发电站
'    ModuleTypes(22) = MoudleTypeCreate("火力发电站", "利用矿物产生电力", 1000, 1, False, -300, 100, 1000, 60)
'    With ModuleTypes(22)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceMetel, 100)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectResource
'        .Effects(1).EffectResources = ResourceCreate(ResourceMineral, -100)
'    End With

    '火力发电站(小)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceMineral, -0.05))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("火力发电站(小)", Effects, Resources, "小型火力发电站。利用矿物产生电力。", 100, 50, False, -1, 100, 0.1, 60)
    
    '火力发电站(中)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectSolarPower, ResourceCreate(ResourceMineral, -50))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("火力发电站(中)", Effects, Resources, "中型火力发电站。利用矿物产生电力。", 100000, 50000, False, -1000, 100000, 100, 120)
    
    '火力发电站(大)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectSolarPower, ResourceCreate(ResourceMineral, -50000))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("火力发电站(大)", Effects, Resources, "大型火力发电站。利用矿物产生电力。", 100000000, 50000000, False, -1000000, 100000000, 100000, 180)

'    '居住模块
'    ModuleTypes(2) = MoudleTypeCreate("居住模块", "提供住宅空间，需要电力维持。", 100, 1, False, -10, 0, -10, 60)
'    With ModuleTypes(2)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceComposites, 10)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectHousing
'        .Effects(1).Amont = 100
'    End With

    '居住模块(小)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(2)
    Resources(1) = ResourceCreate(ResourceComposites, 10)
    Resources(2) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectHousing, ResourceNull, 1000)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("居住模块(小)", Effects, Resources, "小型居住模块。提供住宅空间，需要电力维持。", 100, 0.1, False, -1, 0, -0.02, 60)
    
    '居住模块(中)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(2)
    Resources(1) = ResourceCreate(ResourceComposites, 10000)
    Resources(2) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectHousing, ResourceNull, 1000000)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("居住模块(中)", Effects, Resources, "中型居住模块。提供住宅空间，需要电力维持。", 100000, 100, False, -1000, 0, -20, 120)
    
    '居住模块(大)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(2)
    Resources(1) = ResourceCreate(ResourceComposites, 10000000)
    Resources(2) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectHousing, ResourceNull, 1000000)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("居住模块(大)", Effects, Resources, "大型居住模块。提供住宅空间，需要电力维持。", 100000000, 100000, False, -1000000, 0, -20000, 180)

'    '贸易公司
'    ModuleTypes(4) = MoudleTypeCreate("贸易公司", "产出资金。", 100, 1, False, 50, 10, -10, 60)
'    With ModuleTypes(4)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceComposites, 20)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectNone
'    End With

    '贸易公司(小)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 10)
    ReDim Effects(1)
    Effects(1) = EffectNull
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("贸易公司(小)", Effects, Resources, "小型贸易公司。产出资金。", 100, 0.1, False, 5, 100, -0.01, 60)
    
    '贸易公司(中)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 10000)
    ReDim Effects(1)
    Effects(1) = EffectNull
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("贸易公司(中)", Effects, Resources, "中型贸易公司。产出资金。", 100000, 100, False, 5000, 100000, -10, 120)
    
    '贸易公司(大)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectNull
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("贸易公司(大)", Effects, Resources, "大型贸易公司。产出资金。", 100000000, 100000, False, 5000000, 100000000, -10000, 180)
    
'    '水处理器
'    ModuleTypes(6) = MoudleTypeCreate("水处理器", "产出饮用水。", 100, 1, False, -30, 0, -30, 60)
'    With ModuleTypes(6)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceComposites, 20)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectResource
'        .Effects(1).EffectResources = ResourceCreate(ResourceCleanWater, 30)
'    End With

    '水处理器(小)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 20)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceCleanWater, 0.1))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("水处理器(小)", Effects, Resources, "小型水处理器。产出饮用水。", 50, 0.1, False, -1, 100, -0.02, 60)
    
    '水处理器(中)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 20000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceCleanWater, 100))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("水处理器(中)", Effects, Resources, "中型水处理器。产出饮用水。", 50000, 100, False, -1000, 100000, -20, 120)
    
    '水处理器(大)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 20000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceCleanWater, 100000))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("水处理器(大)", Effects, Resources, "大型水处理器。产出饮用水。", 50000000, 100000, False, -1000000, 100000000, -20000, 180)

'    '矿场
'    ModuleTypes(7) = MoudleTypeCreate("矿场", "产出金属。", 100, 1, False, -10, 10, -30, 60)
'    With ModuleTypes(7)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceMetel, 20)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectResource
'        .Effects(1).EffectResources = ResourceCreate(ResourceMetel, 10)
'    End With

    '矿场(小)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceMetel, 0.1))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("矿场(小)", Effects, Resources, "小型矿场。产出金属。", 20, 10, False, -1, 100, -0.02, 60)
    
    '矿场(中)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceMetel, 100))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("矿场(中)", Effects, Resources, "中型矿场。产出金属。", 20000, 10000, False, -1000, 100000, -20, 120)
    
    '矿场(大)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceMetel, 100000))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("矿场(大)", Effects, Resources, "大型矿场。产出金属。", 20000000, 10000000, False, -1000000, 100000000, -20000, 180)

'    '气体发生器
'    ModuleTypes(8) = MoudleTypeCreate("气体发生器", "向大气中释放气体。", 100, 1, False, -10, 10, -10, 60)
'    With ModuleTypes(8)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceRock, 10)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectGas
'        .Effects(1).Amont = 0.00001
'    End With

    '气体发生器(小)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectGas, ResourceNull, 0.00000001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("气体发生器(小)", Effects, Resources, "小型气体发生器。向大气中释放气体。", 20, 0.1, False, -1, 100, -0.02, 60)
    
    '气体发生器(中)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectGas, ResourceNull, 0.00001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("气体发生器(中)", Effects, Resources, "中型气体发生器。向大气中释放气体。", 20000, 100, False, -1000, 100000, -20, 120)
    
    '气体发生器(大)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectGas, ResourceNull, 0.01)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("气体发生器(大)", Effects, Resources, "大型气体发生器。向大气中释放气体。", 20000000, 100000, False, -1000000, 100000000, -20000, 180)

'    '气体储罐
'    ModuleTypes(9) = MoudleTypeCreate("气体储罐", "吸收大气中的气体。", 100, 1, False, -10, 10, -10, 60)
'    With ModuleTypes(9)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceRock, 10)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectGas
'        .Effects(1).Amont = -0.00001
'    End With

    '气体储罐(小)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectGas, ResourceNull, -0.00000001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("气体储罐(小)", Effects, Resources, "小型气体储罐。吸收大气中的气体。", 20, 0.1, False, -1, 100, -0.02, 60)
    
    '气体储罐(中)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectGas, ResourceNull, -0.00001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("气体储罐(中)", Effects, Resources, "中型气体储罐。吸收大气中的气体。", 20000, 100, False, -1000, 100000, -20, 120)
    
    '气体储罐(大)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectGas, ResourceNull, -0.01)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("气体储罐(大)", Effects, Resources, "大型气体储罐。吸收大气中的气体。", 20000000, 100000, False, -1000000, 100000000, -20000, 180)

'    '地下水矿井
'    ModuleTypes(10) = MoudleTypeCreate("地下水矿井", "升高海平面。", 100, 1, False, -10, 10, -10, 60)
'    With ModuleTypes(10)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceRock, 10)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectWater
'        .Effects(1).Amont = 0.005
'    End With

    '地下水矿井(小)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectWater, ResourceNull, 0.000001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("地下水矿井(小)", Effects, Resources, "小型地下水矿井。升高海平面。", 20, 0.1, False, -1, 100, -0.02, 60)
    
    '地下水矿井(中)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectWater, ResourceNull, 0.001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("地下水矿井(中)", Effects, Resources, "中型地下水矿井。升高海平面。", 20000, 100, False, -1000, 100000, -20, 120)
    
    '地下水矿井(大)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectWater, ResourceNull, 1)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("地下水矿井(大)", Effects, Resources, "大型地下水矿井。升高海平面。", 20000000, 100000, False, -1000000, 100000000, -20000, 180)

'    '水固化系统
'    ModuleTypes(11) = MoudleTypeCreate("水固化系统", "降低海平面。", 100, 1, False, -10, 10, -10, 60)
'    With ModuleTypes(11)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceRock, 10)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectWater
'        .Effects(1).Amont = -0.005
'    End With

    '水固化系统(小)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectWater, ResourceNull, -0.000001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("水固化系统(小)", Effects, Resources, "小型水固化系统。降低海平面。", 20, 0.1, False, -1, 100, -0.02, 60)
    
    '水固化系统(中)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectWater, ResourceNull, -0.001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("水固化系统(中)", Effects, Resources, "中型水固化系统。降低海平面。", 20000, 100, False, -1000, 100000, -20, 120)
    
    '水固化系统(大)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectWater, ResourceNull, -1)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("水固化系统(大)", Effects, Resources, "大型水固化系统。降低海平面。", 20000000, 100000, False, -1000000, 100000000, -20000, 180)

'    '岩石热解系统
'    ModuleTypes(12) = MoudleTypeCreate("岩石热解系统", "向大气层释放氧气。", 100, 1, False, -10, 10, -10, 60)
'    With ModuleTypes(12)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceRock, 20)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectOxygen
'        .Effects(1).Amont = 0.00001
'    End With

    '岩石热解系统(小)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectOxygen, ResourceNull, 0.00000001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("岩石热解系统(小)", Effects, Resources, "小型岩石热解系统。向大气层释放氧气。", 20, 0.1, False, -1, 100, -0.02, 60)
    
    '岩石热解系统(中)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectOxygen, ResourceNull, 0.00001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("岩石热解系统(中)", Effects, Resources, "中型岩石热解系统。向大气层释放氧气。", 20000, 100, False, -1000, 100000, -20, 120)
    
    '岩石热解系统(大)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectOxygen, ResourceNull, 0.01)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("岩石热解系统(大)", Effects, Resources, "大型岩石热解系统。向大气层释放氧气。", 20000000, 100000, False, -1000000, 100000000, -20000, 180)

'    '氧气固化器
'    ModuleTypes(13) = MoudleTypeCreate("氧气固化器", "除去大气层中的氧气。", 100, 1, False, -10, 10, -10, 60)
'    With ModuleTypes(13)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceRock, 10)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectOxygen
'        .Effects(1).Amont = -0.00001
'    End With

    '氧气吸收器(小)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectOxygen, ResourceNull, 0.00000001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("氧气吸收器(小)", Effects, Resources, "小型氧气吸收器。除去大气层中的氧气。", 20, 0.1, False, -1, 100, -0.02, 60)
    
    '氧气吸收器(中)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectOxygen, ResourceNull, 0.00001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("氧气吸收器(中)", Effects, Resources, "中型氧气吸收器。除去大气层中的氧气。", 20000, 100, False, -1000, 100000, -20, 120)
    
    '氧气吸收器(大)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectOxygen, ResourceNull, 0.01)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("氧气吸收器(大)", Effects, Resources, "大型氧气吸收器。除去大气层中的氧气。", 20000000, 100000, False, -1000000, 100000000, -20000, 180)

'    '行星加热器
'    ModuleTypes(14) = MoudleTypeCreate("行星加热器", "升高行星温度。", 100, 1, False, -10, 10, -10, 60)
'    With ModuleTypes(14)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceRock, 10)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectTempreture
'        .Effects(1).Amont = 0.5
'    End With

    '行星加热器(小)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectTempreture, ResourceNull, 0.0001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("行星加热器(小)", Effects, Resources, "小型行星加热器。降低海平面。", 20, 0.1, False, -1, 100, -0.02, 60)
    
    '行星加热器(中)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectWater, ResourceNull, 0.1)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("行星加热器(中)", Effects, Resources, "中型行星加热器。降低海平面。", 20000, 100, False, -1000, 100000, -20, 120)
    
    '行星加热器(大)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectWater, ResourceNull, 100)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("行星加热器(大)", Effects, Resources, "大型行星加热器。降低海平面。", 20000000, 100000, False, -1000000, 100000000, -20000, 180)

'    '机器人工厂
'    ModuleTypes(15) = MoudleTypeCreate("机器人工厂", "使用金属制造机器人。", 100, 1, False, -10, 10, -30, 60)
'    With ModuleTypes(15)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceMetel, 30)
'        ReDim .Effects(2)
'        .Effects(1).Type = EffectResource
'        .Effects(1).EffectResources = ResourceCreate(ResourceMetel, -10)
'        .Effects(2).Type = EffectResource
'        .Effects(2).EffectResources = ResourceCreate(ResourceRobot, 10)
'    End With
'
'    '研究所
'    ModuleTypes(16) = MoudleTypeCreate("研究所", "进行研究。", 100, 1, False, -10, 10, -10, 60)
'    With ModuleTypes(16)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceComposites, 20)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectResearchPoint
'        .Effects(1).Amont = 1
'    End With
'
'    '电视台
'    ModuleTypes(17) = MoudleTypeCreate("电视台", "增加声望。", 100, 1, False, -10, 10, -10, 60)
'    With ModuleTypes(17)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceComposites, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectPrestige
'        .Effects(1).Amont = 1
'    End With
'
'    '仓库
'    ModuleTypes(18) = MoudleTypeCreate("仓库", "增加物品储量。", 100, 1, False, 0, 10, 0, 60)
'    With ModuleTypes(18)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectStorage
'        .Effects(1).Amont = 500
'    End With
'
'    '自来水厂
'    ModuleTypes(19) = MoudleTypeCreate("自来水厂", "从自然水体中获取水并生产净水。", 10000000, 10, False, -1000000, 10000000, -10000000, 300)
'    With ModuleTypes(19)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceComposites, 20000)
'        ReDim .Effects(2)
'        .Effects(1).Type = EffectResource
'        .Effects(1).EffectResources = ResourceCreate(ResourceCleanWater, 1000000000)
'        .Effects(2).Type = EffectRunOff
'        .Effects(2).Amont = -1000000000
'    End With
'
'    '旅行社
'    ModuleTypes(20) = MoudleTypeCreate("旅行社", "解锁旅游业(未完成)。", 100, 1, False, -30, 10, -30, 60)
'    With ModuleTypes(20)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceComposites, 20)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectResource
'        .Effects(1).EffectResources = ResourceCreate(ResourceCleanWater, 10)
'    End With
'
'    '传统住房
'    ModuleTypes(23) = MoudleTypeCreate("传统住房", "提供住房", 10000000, 10, True, -10000000, 10000000, -10000000, 300)
'    With ModuleTypes(23)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceMetel, 20000)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectHousing
'        .Effects(1).Amont = 1000000000
'    End With
'
'    '反光板阵列
'    ModuleTypes(24) = MoudleTypeCreate("反光板阵列", "增加星球的反射率", 100, 1, False, -10, 0, 0, 60)
'    With ModuleTypes(24)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceComposites, 10)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectReflectivity
'        .Effects(1).Amont = 0.99
'    End With
'
'    '火力发电区划
'    ModuleTypes(25) = MoudleTypeCreate("火力发电区划", "使用燃料发电", 1000, 1, True, -3000000, 1000000, 10000000, 300)
'    With ModuleTypes(25)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceComposites, 10000)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10000)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectResource
'        .Effects(1).EffectResources = ResourceCreate(ResourceMineral, -1000000)
'    End With
'
'    '传统矿场
'    ModuleTypes(26) = MoudleTypeCreate("传统矿场", "产出矿物。", 10000000, 1, False, -1000000, 50000000, -10000000, 300)
'    With ModuleTypes(26)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceMetel, 20000)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectResource
'        .Effects(1).EffectResources = ResourceCreate(ResourceMineral, 10000000)
'    End With
End Sub

                    
'加载航天器
Private Sub LoadSpaceCraft()
    Dim SpacePosition As SpacePosition
    Dim Effects() As Effect
    Dim Resources() As Resource
    
    ReDim Spacecrafts(1)
    With Spacecrafts(1)
        .Name = "希望号(殖民船)"
        .Maintenance = 1
        .Population = 100
        .Power = 1
        .Position = SpacePositionCreate(SpacePositionTypeLand, 3)
        .Construction = 1
        .Space = 10
        .Enabled = True
        
        ReDim Resources(2)
        Resources(1) = ResourceCreate(ResourceComposites, 200)
        Resources(2) = ResourceCreate(ResourceMetel, 200)
        
        .Storage = Resources
        
        ReDim Effects(5)
        Effects(1) = EffectCreate(EffectHousing, ResourceNull, 150)
        Effects(2) = EffectCreate(EffectStorage, ResourceNull, 500)
        Effects(3) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 0.015))
        Effects(4) = EffectCreate(EffectResource, ResourceCreate(ResourceCleanWater, 0.015))
        
        .Effects = Effects
    End With
End Sub

Private Sub LoadSystem(System As System)
    Dim i As Long
    
    ReDim Stars(1)
    With Stars(1)
        .Name = "太阳"
        .Tag = "SUN"
        .Color = vbYellow
        .Magnitude = 4.83
        .Mass = 198900000000# '亿亿吨
        .Type = StarTypeMainSequence
    End With
    
    ReDim Planets(4)
    '水星
    With Planets(1)
        .Name = "水星"
        .Tag = "SUN1"
        .Color = RGB(192, 192, 192)
        .Radio = 2440 '千米
'            .Water = 0 '亿亿吨
'            .Mass = 33011  '亿亿吨
        .Tempreture = 452 '开尔文
'            .Oxygen = 0 '亿亿吨
        .Reflectivity = 0.119
        .RotationPeriod = 58.646 '天
        .OrbitRadius = 0.5791 '亿千米
        ReDim .Materials(2)
        .Materials(1).Type = "金属"
        .Materials(1).Mass = 17811
        .Materials(2).Type = "岩石"
        .Materials(2).Mass = 15200
    End With
    '金星
    With Planets(2)
        .Name = "金星"
        .Tag = "SUN2"
        .Color = RGB(255, 192, 72)
        .Radio = 6052 '千米
'            .Water = 0 '亿亿吨
'            .Mass = 486750  '亿亿吨
        .Tempreture = 737 '开尔文
'            .Oxygen = 0 '亿亿吨
        .Reflectivity = 0.75
'            .Gas = 48.27 '亿亿吨
        .RotationPeriod = 243 '天
        .OrbitRadius = 1.082 '亿千米
        ReDim .Materials(4)
        .Materials(1).Type = "金属"
        .Materials(1).Mass = 97750
        .Materials(2).Type = "岩石"
        .Materials(2).Mass = 389000
        .Materials(3).Type = "碳氧化物气体"
        .Materials(3).Mass = 46.82
        .Materials(4).Type = "稳定气体"
        .Materials(4).Mass = 1.45
    End With
    '地球
    With Planets(3)
        .Name = "地球"
        .Tag = "SUN3"
        .Color = vbBlue
        .Radio = 6371 '千米
'            .Water = 136 '亿亿吨
'            .Mass = 597237  '亿亿吨
        .Tempreture = 289 '开尔文
'            .Oxygen = 0.11885 '亿亿吨
        .Reflectivity = 0.29
'            .Gas = 0.5136 - 0.11885 '亿亿吨
        .RotationPeriod = 1 '天
        .OrbitRadius = 1.496 '亿千米
        ReDim .Materials(5)
        .Materials(1).Type = "金属"
        .Materials(1).Mass = 147237
        .Materials(2).Type = "岩石"
        .Materials(2).Mass = 450000
        .Materials(3).Type = "水"
        .Materials(3).Mass = 136
        .Materials(4).Type = "氧气"
        .Materials(4).Mass = 0.11885
        .Materials(5).Type = "稳定气体"
        .Materials(5).Mass = 0.39475
        
        .HomeWorld = True
        .Colonys.Population = 7000000000#
        
    End With
    '火星
    With Planets(4)
        .Name = "火星"
        .Tag = "SUN4"
        .Color = vbRed
        .Radio = 3389 '千米
'            .Water = 0.0000001 '亿亿吨
'            .Mass = 64171  '亿亿吨
        .Tempreture = 218 '开尔文
'            .Oxygen = 0 '亿亿吨
        .Reflectivity = 0.16
'            .Gas = 0.0025 '亿亿吨
        .RotationPeriod = 1.0259 '天
        .OrbitRadius = 2.279 '亿千米
        ReDim .Materials(5)
        .Materials(1).Type = "金属"
        .Materials(1).Mass = 6471
        .Materials(2).Type = "岩石"
        .Materials(2).Mass = 57700
        .Materials(3).Type = "冰"
        .Materials(3).Mass = 0.0000001
        .Materials(4).Type = "碳氧化物气体"
        .Materials(4).Mass = 0.002375
        .Materials(5).Type = "稳定气体"
        .Materials(5).Mass = 0.000125
    End With
        
    For i = 1 To UBound(Planets)
        With Planets(i)
            ReDim .Resources(ResourceCount - 1)
            ReDim .Market.Prices(ResourceCount - 1)
            ReDim .Market.Storage(ResourceCount - 1)
            ReDim .Market.Money(ResourceCount - 1)
            ReDim .Colonys.PopulationStorage(ResourceCount - 1)
            ReDim .Transport(ResourceCount - 1)
            ReDim .Modules(0)
            If i = 3 Then
                .UtilizingBlock = Int(0.3 * PlanetGetBlock(Planets(i)) + 3)
                PlanetAddMoudle Planets(i), 19, 7, 1
                PlanetAddMoudle Planets(i), 21, 70, 1
                PlanetAddMoudle Planets(i), 23, 7, 1
                PlanetAddMoudle Planets(i), 25, 94, 1
                PlanetAddMoudle Planets(i), 26, 10, 1
                .Resources(ResourceMineral) = 10000000
            Else
                .UtilizingBlock = Int(0.01 * PlanetGetBlock(Planets(i)) + 3)
            End If
        End With
    Next
    
    With System
        .Name = "太阳系"
        ReDim .Stars(1)
        .Stars(1) = "SUN"
        
        ReDim .Planets(4)
        .Planets(1) = "SUN1" '水星
        .Planets(2) = "SUN2" '金星
        .Planets(3) = "SUN3" '地球
        .Planets(4) = "SUN4" '火星
    End With
End Sub

Private Sub LoadTechnology()
    ReDim Technologys(5)
    
    '核聚变
    With Technologys(1)
        .Name = "核聚变"
        .NeedPoints = 1000
        .IsResearched = False
    End With
    
    '人工智能
    With Technologys(2)
        .Name = "人工智能"
        .NeedPoints = 100
        .IsResearched = False
    End With
    
    '亚光速飞船
    With Technologys(3)
        .Name = "亚光速飞船"
        .NeedPoints = 200
        .IsResearched = False
    End With
    
    '巨型建筑
    With Technologys(4)
        .Name = "巨型建筑"
        .NeedPoints = 2000
        .IsResearched = False
    End With
    
    '医学
    With Technologys(5)
        .Name = "医学"
        .NeedPoints = 100
        .IsResearched = False
    End With
    
    '催化
    With Technologys(6)
        .Name = "催化"
        .NeedPoints = 100
        .IsResearched = False
    End With
    
    '无线输电
    With Technologys(6)
        .Name = "无线输电"
        .NeedPoints = 100
        .IsResearched = False
    End With
    
    '虚拟现实
    With Technologys(6)
        .Name = "虚拟现实"
        .NeedPoints = 100
        .IsResearched = False
    End With
End Sub

''计算市场价格
'Private Sub MarketCalculatePrice(Market As Market)
'
'    Market
'    MarketCalculatePrice = Market.Money * Market.Prices
'End Sub

'获取市场乘积
Private Function MarketGetProduct(Market As Market, Resource As ResourceEnum) As Double
    MarketGetProduct = Market.Money(Resource) * Market.Storage(Resource)
End Function

Private Function MaterialCreate(ByVal MaterialType As String, ByVal Mass As Double) As Material
    With MaterialCreate
        .Type = MaterialType
        .Mass = Mass
    End With
End Function

Private Function MaterialGetState(Material As Material) As MaterialStateEnum
    MaterialGetState = MaterialGetType(Material).State
End Function

Private Function MaterialGetType(Material As Material) As MaterialType
    Dim i As Long
    
    For i = 1 To UBound(MaterialTypes)
        With MaterialTypes(i)
            If .Name = Material.Type Then
                MaterialGetType = MaterialTypes(i)
                Exit Function
            End If
        End With
    Next
End Function

Private Function Max(ByVal a As Variant, ByVal b As Variant) As Variant
    If a > b Then
        Max = a
    Else
        Max = b
    End If
End Function

Private Function Min(ByVal a As Variant, ByVal b As Variant) As Variant
    If a < b Then
        Min = a
    Else
        Min = b
    End If
End Function

Private Sub ModuleDrawUI(Module As Module, ByVal X As Long, ByVal Y As Long)
    Dim DrawText As String
    Dim i As Long, j As Long
    
    With Module
        If ModuleTypes(.Type).Name = "" Then Exit Sub
        
        FillColor = vbWhite
        Line (X, Y)-(X + 200, Y + 300), , B
        
        MyPrint ModuleTypes(.Type).Name, X + 10, Y + 10
        MyPrint "规模:" & .Size, X + 10, CurrentY
        MyPrint "能源:" & ModuleTypes(.Type).Power * .Size, X + 10, CurrentY
        MyPrint "资金:" & ModuleTypes(.Type).Maintenance * .Size, X + 10, CurrentY
        MyPrint "岗位:" & ModuleTypes(.Type).Staff * .Size, X + 10, CurrentY
        
        For i = 1 To UBound(ModuleTypes(.Type).Effects)
            With ModuleTypes(.Type).Effects(i)
                Select Case .Type
                Case EffectNone
                    MyPrint GetEffectName(.Type), X + 10, CurrentY
                Case EffectResource
                    MyPrint GetResourceName(.EffectResources.Type) & ":" & .EffectResources.Amont * Module.Size, X + 10, CurrentY
                Case Else
                    MyPrint GetEffectName(.Type) & ":" & .Amont * Module.Size, X + 10, CurrentY
                End Select
            End With
        Next
        
        MyPrint "占用空间:" & ModuleTypes(.Type).Space * .Size, X + 10, CurrentY
        MyPrint "效率:" & Format(.Efficiency, "0.00%"), X + 10, CurrentY
        
        If .Construction < 1 Then
            MyPrint "建造中", X + 10, CurrentY
            DrawProgressBar RectangleCreate(X + 10, CurrentY + 2, 100, 14), .Construction, RGB(120, 255, 120)
        End If
        
        MyPrint "最高承受温度" & Format(ModuleTypes(.Type).MaxTempreture - 273.15, "0.00") & "℃", X + 10, CurrentY + 2
        MyPrint "最大承受压力" & Format(ModuleTypes(.Type).MaxPressure / 1000000, "0.00") & "MPa", X + 10, CurrentY
        
        MyPrint "仓储:", X + 10, CurrentY
        
        If UBound(.Storage) = 0 Then
            MyPrint "无", X + 10, CurrentY
        Else
            For i = 1 To UBound(.Storage)
                MyPrint ResourceToString(.Storage(i)), X + 10, CurrentY
            Next
        End If
        
        If Not Planets(SelectPlanet).HomeWorld Then
            DrawButtonWithUI RectangleCreate(X + 10, Y + 220, 60, 30), "拆除", "system(""dismantle_module " & Planets(SelectPlanet).Tag & " " & SelectModule & """)"
            
            If Module.Construction = 1 Then
                '绘制启/禁用按钮
                If Module.Enabled Then
                    DrawText = "禁用"
                Else
                    DrawText = "启用"
                End If
                DrawButtonWithUI RectangleCreate(X + 80, Y + 220, 60, 30), DrawText, "system(""switch_module_enabled " & SelectPlanet & " " & SelectModule & """)"
                
                DrawButtonWithUI RectangleCreate(X + 10, Y + 260, 60, 30), "扩建", "system(""expand_module " & SelectPlanet & " " & SelectModule & """)" '绘制扩建按钮
            End If
        End If
    End With
    
    DrawCloseButton X + 170, Y + 5, "system(""clear_select"")"
End Sub

'模块效果计算
Private Sub MoudleEffectCalculate(Module As Module, Planet As Planet)
    Dim GetModuleType As ModuleType
    Dim i As Long
    
    With Module
        GetModuleType = ModuleTypes(.Type)
        For i = 1 To UBound(GetModuleType.Effects)
            EffectCalculate GetModuleType.Effects(i), Planet, .Size * .Efficiency
        Next
    End With
End Sub

'模块效率计算
Private Sub MoudleEfficiencyCalculate(Module As Module, Planet As Planet, ByVal ID As Long)
    Dim GetModuleType As ModuleType
    Dim GetResources As Resource
    Dim i As Long
    
    With Module
        '计算效率修正
        .EfficiencyModifier = 1
        
        If .Enabled = False Then
            .EfficiencyModifier = 0
        End If
        
        '太阳能效率修正
        For i = 1 To UBound(ModuleTypes(.Type).Effects)
            If ModuleTypes(.Type).Effects(i).Type = EffectSolarPower Then
                .EfficiencyModifier = .EfficiencyModifier * PlanetGetSolarPower(Stars(1), Planet)
            End If
        Next
        
        If .EfficiencyModifier < 0 Then
            .EfficiencyModifier = 0
        End If
        .Efficiency = .EfficiencyModifier
        
        '效率修正计算完毕，计算效率
        '电力
        If ModuleTypes(.Type).Power < 0 Then
            If .Efficiency > .EfficiencyModifier * GetPowerAdyquacy(Planet) Then
                .Efficiency = .EfficiencyModifier * GetPowerAdyquacy(Planet)
            End If
        End If
        
        '人口
        If ModuleTypes(.Type).Staff > 0 Then
            If .Efficiency > .EfficiencyModifier * GetStaffAdyquacy(Planet) Then
                .Efficiency = .EfficiencyModifier * GetStaffAdyquacy(Planet)
            End If
        End If
        
        '资源效率修正
        GetModuleType = ModuleTypes(.Type)
        For i = 1 To UBound(GetModuleType.Effects)
            Select Case GetModuleType.Effects(i).Type
            Case EffectResource
                GetResources = GetModuleType.Effects(i).EffectResources
                If GetResources.Amont < 0 Then
                    If .Efficiency > .EfficiencyModifier * (Planet.Resources(GetResources.Type) / Abs(GetResources.Amont)) Then
                        .Efficiency = .EfficiencyModifier * (Planet.Resources(GetResources.Type) / Abs(GetResources.Amont))
                    End If
                End If
            End Select
        Next
        
        If .Efficiency < 0 Then
            .Efficiency = 0
        End If
    End With
End Sub

Private Function MoudleTypeNull() As ModuleType
    Dim Effects(0) As Effect
    Dim Resources(0) As Resource
    MoudleTypeNull = MoudleTypeCreate("EMPTY_MODULE", Effects, Resources)
End Function

Private Function MoudleTypeCreate(ByVal Name As String, Effects() As Effect, Resources() As Resource, Optional ByVal Description As String = "", Optional ByVal Cost As Double = 0, Optional ByVal Space As Double = 0, Optional ByVal LivableRequire As Boolean = False, Optional ByVal Maintenance As Double = 0, Optional ByVal Staff As Long = 0, Optional ByVal Power As Double = 0, Optional ByVal BuildTime As Double = 0, Optional ByVal MaxTempreture As Double = 2400, Optional ByVal MaxPressure As Double = 40000000) As ModuleType
    With MoudleTypeCreate
        .Name = Name
        .Description = Description
        .Cost = Cost
        .Effects = Effects
        .Resources = Resources
        .Space = Space
        .LivableRequire = LivableRequire
        .Maintenance = Maintenance
        .Staff = Staff
        .Power = Power
        .BuildTime = BuildTime
        .MaxTempreture = MaxTempreture
        .MaxPressure = MaxPressure
    End With
End Function

Private Function MousePosition() As PointApi '获取鼠标位置
    Dim CurPos As Long
    Dim PixelLeft As Long
    Dim PixelTop As Long
    Dim CursorPosition As PointApi
    
    PixelLeft = Left / GetDpi + GetFrameWidth
    PixelTop = Top / GetDpi + GetFrameTop
    CurPos = GetCursorPos(CursorPosition)
    MousePosition.X = CursorPosition.X - PixelLeft
    MousePosition.Y = CursorPosition.Y - PixelTop
End Function

Private Sub MyPrint(ByVal Text As String, ByVal X As Long, ByVal Y As Long, Optional ByVal Mode As Long)
    Dim Lines() As String
    Dim i As Long
    
    'Mode0为输入打印左上角位置
    'Mode1为输入打印中心位置
    If Mode = 1 Then
        X = X - 0.5 * TextWidth(Text)
        Y = Y - 0.5 * TextHeight(Text)
    End If
    
    '按行分割，每一行都先设置缩进再打印
    CurrentY = Y
    Lines = Split(Text, vbCrLf)
    For i = LBound(Lines) To UBound(Lines)
        CurrentX = X
        Print Lines(i)
    Next
End Sub

'数字格式化函数
Private Function NumberFormat(n As Variant) As String
    If Abs(n) <= 10000 Then
        NumberFormat = Format(n, "0")
    Else
        If Abs(n) <= 100000000 Then
            NumberFormat = Format(n / 10000, "0.0") & "万"
        Else
            NumberFormat = Format(n / 100000000, "0.0") & "亿"
        End If
    End If
End Function

Private Sub PlanetAddMoudle(Planet As Planet, ByVal MoudleType As Long, Optional ByVal Size As Long = 0, Optional ByVal Construction As Double = 0, Optional ByVal Enabled As Boolean = True)
    ReDim Preserve Planet.Modules(UBound(Planet.Modules) + 1)
    With Planet.Modules(UBound(Planet.Modules))
        .Type = MoudleType
        .Size = Size
        .Construction = Construction
        .Enabled = Enabled
        ReDim .Storage(0)
    End With
End Sub

'行星添加资源
Private Sub PlanetAddResource(Planet As Planet, ResourceType As ResourceEnum, Amont As Double)
    Planet.Resources(ResourceType) = Planet.Resources(ResourceType) + Amont
    If Planet.Resources(ResourceType) < 0 Then Planet.Resources(ResourceType) = 0
End Sub

Private Sub PlanetBuildModule(Planet As Planet, ByVal MoudleType As Long) '建造建筑
    Dim i As Long
    
    '检查是否建设在母星
    If Planet.HomeWorld Then
        MsgBox "建筑无法建设在母星"
        Exit Sub
    End If
    
    With ModuleTypes(MoudleType)
        '检查建筑宜居条件是否满足
        If Not PlanetIsLiveable(Planet) And .LivableRequire Then
            MsgBox "建造该建筑需要星球宜居"
            Exit Sub
        End If
    
        '检查资金是否足够
        If Money < .Cost Then
            MsgBox "缺少" & NumberFormat(.Cost - Money) & "资金"
            Exit Sub
        End If
        
        '检查空间是否足够
        If .Space > 0 Then
            If PlanetGetUsedBlock(Planet) + .Space > Planet.UtilizingBlock Then
                MsgBox Planet.Name & "缺少" & NumberFormat(PlanetGetUsedBlock(Planet) + .Space - Planet.UtilizingBlock) & "建筑空间"
                Exit Sub
            End If
        End If
        
        '检查资源是否足够
        For i = 1 To UBound(.Resources)
            If Planet.Resources(.Resources(i).Type) < .Resources(i).Amont Then
                MsgBox "缺少" & NumberFormat(.Resources(i).Amont - Planet.Resources(.Resources(i).Type)) & GetResourceName(.Resources(i).Type)
                Exit Sub
            End If
        Next
    
        PlanetAddMoudle Planet, MoudleType
    
        '扣除资源
        Money = Money - .Cost
        For i = 1 To UBound(.Resources)
            PlanetAddResource Planet, .Resources(i).Type, -.Resources(i).Amont
        Next
    End With
End Sub

'改变星球上的某种物质含量
Private Sub PlanetChangeMaterial(Planet As Planet, ByVal MaterialType As String, ByVal Mass As Double)
    Dim i As Long
    With Planet
        For i = 1 To UBound(.Materials)
            With .Materials(i)
                If .Type = MaterialType Then
                    .Mass = .Mass + Mass
                    If .Mass <= 0 Then PlanetDeleteMaterial Planet, .Type
                    Exit Sub
                End If
            End With
        Next
        If Mass > 0 Then
            ReDim Preserve .Materials(UBound(.Materials) + 1)
            .Materials(UBound(.Materials)) = MaterialCreate(MaterialType, Mass)
        End If
    End With
End Sub

'删除星球上所有某种物质
Private Sub PlanetDeleteMaterial(Planet As Planet, ByVal MaterialType As String)
    Dim i As Long, j As Long
    With Planet
        For i = UBound(.Materials) To 1 Step -1
            If .Materials(i).Type = MaterialType Then
                For j = i To UBound(.Materials) - 1
                    .Materials(j) = .Materials(j + 1)
                Next
                ReDim Preserve .Materials(UBound(.Materials) - 1)
                Exit Sub
            End If
        Next
    End With
End Sub

'删除模块
Private Sub PlanetDeleteMoudle(Planet As Planet, ID As Long)
    Dim i As Long
    For i = ID To UBound(Planet.Modules) - 1
        Planet.Modules(i) = Planet.Modules(i + 1)
    Next
    ReDim Preserve Planet.Modules(UBound(Planet.Modules) - 1)
End Sub

'绘制行星
Private Sub PlanetDraw(Planet As Planet, X As Long, Y As Long)
    With Planet
        FillColor = .Color
        Circle (X, Y), 5
        CurrentX = X - 0.5 * TextWidth(.Name)
        CurrentY = Y + 10
        Print .Name
        MyPrint .Name, X - 0.5 * TextWidth(.Name), Y + 10
    End With
End Sub

Private Function PlanetGetBlock(Planet As Planet) As Long
    PlanetGetBlock = Int(0.2 * Planet.Radio ^ 1.1) + 10
End Function

'获取行星蒸发量
Private Function PlanetGetEvaporation(Planet As Planet) As Double
    PlanetGetEvaporation = 200000000000# * (Planet.Tempreture / 289) * (Min(PlanetGetWater(Planet), 1) / 0.718) '* (PlanetGetSolarPower(Star, Planet) / 0.72)
End Function

'获取行星温室效应
Private Function PlanetGetGreenhouseEffect(Planet As Planet) As Double
    PlanetGetGreenhouseEffect = Tanh(0.41 * Log(PlanetGetPressure(Planet) / 100000 + 1))
End Function

'获取行星的ID值
Private Function PlanetGetID(Planet As Planet) As Long
    PlanetGetID = FindPlanetWithTag(Planet.Tag)
End Function

Private Function PlanetGetSystem(Planet As Planet) As Long
    Dim PlanetID As Long
    
    PlanetID = PlanetGetID(Planet)
'    Dim i As Long
'    For i = 1 To UBound(Planets)
'        If Planets(i).Tag = Tag Then
'            FindPlanetWithTag = i
'        End If
'    Next
    
End Function

Private Function PlanetGetMass(Planet As Planet, Optional ByVal State As MaterialStateEnum = -1, Optional ByVal Material As String = "") As Double
    Dim i As Long
    
    PlanetGetMass = 0
    With Planet
        For i = 1 To UBound(.Materials)
            If Material = "" Then
                If State >= 0 Then
                    If MaterialGetState(Planet.Materials(i)) = State Then PlanetGetMass = PlanetGetMass + Planet.Materials(i).Mass
                Else
                    PlanetGetMass = PlanetGetMass + Planet.Materials(i).Mass
                End If
            Else
                If Planet.Materials(i).Type = Material Then PlanetGetMass = PlanetGetMass + Planet.Materials(i).Mass
            End If
        Next
    End With
End Function

Private Function PlanetGetOrbitalPeriod(Star As Star, Planet As Planet) As Double
    With Planet
        PlanetGetOrbitalPeriod = 2 * 4 * Atn(1) * ((.OrbitRadius * 10 ^ 11) ^ 3 / (6.67259 * 10 ^ -11 * Star.Mass * 10 ^ 19)) ^ 0.5 / 86400
    End With
End Function

Private Function PlanetGetOxygen(Planet As Planet) As Double
    Dim OxygenMass As Double
    Dim i As Long
    
    With Planet
        For i = 1 To UBound(.Materials)
            If Planet.Materials(i).Type = "氧气" Then OxygenMass = OxygenMass + Planet.Materials(i).Mass
        Next
    End With
    
    If OxygenMass = 0 Then
        PlanetGetOxygen = 0
    Else
        PlanetGetOxygen = OxygenMass / PlanetGetMass(Planet, MaterialStateGas)
    End If
End Function

Private Function PlanetGetPressure(Planet As Planet) As Double
    PlanetGetPressure = PlanetGetMass(Planet, MaterialStateGas) / (4 * 4 * Atn(1) * Planet.Radio ^ 2) * 10 ^ 13 * GetGravity(Planet)
End Function

Private Function PlanetGetRainfallDensity(Planet As Planet) As Double
    PlanetGetRainfallDensity = (Planet.Tempreture / 289) * (Min(PlanetGetWater(Planet), 1) / 0.718)
End Function

Private Function PlanetGetRunoff(Planet As Planet) As Double
    PlanetGetRunoff = PlanetGetEvaporation(Planet) * (1 - Min(PlanetGetWater(Planet), 1))
End Function

Private Function PlanetGetSolarPower(Star As Star, Planet As Planet) As Double
    PlanetGetSolarPower = StarGetRelativeLuminosity(Star) / (Planet.OrbitRadius / 1.496) ^ 2 * (1 - Tanh(0.41 * Log(PlanetGetPressure(Planet) / 100000 + 1)))
End Function

Private Function PlanetGetUsedBlock(Planet As Planet) As Long
    Dim i As Long
    
    PlanetGetUsedBlock = 0
    With Planet
        For i = 1 To UBound(.Modules)
            PlanetGetUsedBlock = PlanetGetUsedBlock + ModuleTypes(.Modules(i).Type).Space * .Modules(i).Size
        Next
    End With
End Function

Private Function PlanetGetWater(Planet As Planet) As Double
    Dim WaterMass As Double
    Dim i As Long
    
    With Planet
        For i = 1 To UBound(.Materials)
            If Planet.Materials(i).Type = "水" Then WaterMass = WaterMass + Planet.Materials(i).Mass
        Next
    End With
    PlanetGetWater = WaterMass / 136 * 0.718
End Function

'行星是否宜居
Private Function PlanetIsLiveable(Planet As Planet) As Boolean
    PlanetIsLiveable = True
    
    '检查各项数据，判断是否宜居
    If PlanetGetOxygen(Planet) < 0.18 Or PlanetGetOxygen(Planet) > 0.24 Then PlanetIsLiveable = False
    If PlanetGetWater(Planet) < 0.25 Or PlanetGetWater(Planet) > 0.75 Then PlanetIsLiveable = False
    If Planet.Tempreture < 237 Or Planet.Tempreture > 337 Then PlanetIsLiveable = False
    If PlanetGetPressure(Planet) < 50000 Or PlanetGetPressure(Planet) > 150000 Then PlanetIsLiveable = False
End Function

''人民付出货币，购买物资
'Private Sub PoplationBought(Colony As Colony, ByVal Amont As Double, ByVal Price As Double)
'    Colony.PopulationMoney = Colony.PopulationMoney - Amont * Price
'    Colonys.PopulationStorage(
'                            .Colonys.PopulationMoney = .Colonys.PopulationMoney - MarketActualBought * MarketTempPrice
'                            .Market.Money(ResourceFood) = .Market.Money(ResourceFood) + MarketActualBought * MarketTempPrice
'                            .Market.Storage(ResourceFood) = .Market.Storage(ResourceFood) - MarketActualBought
'                            .Colonys.PopulationStorage(ResourceFood) = .Colonys.PopulationStorage(ResourceFood) + MarketActualBought
'End Sub

Private Function RandomInt(ByVal X As Double) As Long
    RandomInt = Int(X)
    If Rnd + X - Int(X) > 1 Then RandomInt = RandomInt + 1
End Function

'矩形平移
Private Function RectangleTranslation(Rect As Rectangle, Optional ByVal X As Long = 0, Optional ByVal Y As Long = 0) As Rectangle
    With RectangleTranslation
        .Left = Rect.Left + X
        .Top = Rect.Top + Y
        .Width = Rect.Width + X
        .Height = Rect.Height + Y
    End With
End Function

'绘制矩形
Private Sub RectangleDraw(Rectangle As Rectangle, Optional ByVal RectangleFillColor = vbWhite)
    FillColor = RectangleFillColor
    Line (Rectangle.Left, Rectangle.Top)-(Rectangle.Left + Rectangle.Width - 1, Rectangle.Top + Rectangle.Height - 1), , B
End Sub

'判断给定点是否在矩形的真实坐标内(含边界)
Private Function RectangleIsIn(Rect As Rectangle, ByVal X As Long, ByVal Y As Long) As Boolean
    With Rect
        RectangleIsIn = X >= .Left And X <= .Left + .Width And Y >= .Top And Y <= .Top + .Height
    End With
End Function

'创建矩形
Private Function RectangleCreate(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long) As Rectangle
    With RectangleCreate
        .Left = Left
        .Top = Top
        .Width = Width
        .Height = Height
    End With
End Function

'创建资源
Private Function ResourceCreate(ByVal ResourceType As ResourceEnum, ByVal Amont As Double) As Resource
    With ResourceCreate
        .Type = ResourceType
        .Amont = Amont
    End With
End Function

'创建空资源
Private Function ResourceNull() As Resource
    With ResourceNull
        .Type = ResourceNone
        .Amont = 0
    End With
End Function

'以“资源名称：资源数量”形式返回字符串
Private Function ResourceToString(Resource As Resource) As String
    ResourceToString = GetResourceName(Resource.Type) & ":" & NumberFormat(Resource.Amont)
End Function

Private Function SaveGame() As Long '存储游戏
    If Dir(App.Path & "\file\") = "" Then
        MkDir App.Path & "\file"
    End If
    
    Open App.Path & "\file\save.txt" For Output As #1
        Print #1, "systems = {"
        Print #1, "}"
        Print #1, "planets = {"
        Print #1, "}"
    Close
    
    MsgBox "保存成功！"
End Function

'Private Function LoadGame() As Long '加载游戏
'    If Dir(App.Path & "\file\") = "" Then
'        MkDir App.Path & "\file"
'    End If
'
'    Open App.Path & "\file\save.txt" For Output As #1
'
'    Close
'End Function

Private Function SaveLog() As Long '存储游戏
    If Dir(App.Path & "\file\") = "" Then
        MkDir App.Path & "\file"
    End If
    
    Open App.Path & "\file\log.txt" For Output As #1
        Print #1, GameLog
    Close
End Function

'删除飞船上的物资
Private Sub SpacecraftDeleteStorage(Spacecraft As Spacecraft, ID As Long)
    Dim i As Long
    
    For i = ID To UBound(Spacecraft.Storage) - 1
        Spacecraft.Storage(i) = Spacecraft.Storage(i + 1)
    Next
    ReDim Preserve Spacecraft.Storage(UBound(Spacecraft.Storage) - 1)
End Sub

'获取飞船所在的星球
Private Function SpacecraftGetPlanet(Spacecraft As Spacecraft) As Long
    SpacecraftGetPlanet = SpacePositionGetPlanet(Spacecraft.Position)
End Function

'卸载飞船上的物资
Private Sub SpacecraftUnloadStorage(Spacecraft As Spacecraft, ID As Long, Amont As Double)
    Dim i As Long
    
    With Spacecraft
        If .Position.Type = SpacePositionTypeLand Then
            If .Storage(ID).Amont > Amont Then
                .Storage(ID).Amont = .Storage(ID).Amont - Amont
                PlanetAddResource Planets(SpacecraftGetPlanet(Spacecraft)), .Storage(ID).Type, Amont
            Else
                PlanetAddResource Planets(SpacecraftGetPlanet(Spacecraft)), .Storage(ID).Type, .Storage(ID).Amont
                SpacecraftDeleteStorage Spacecraft, ID
            End If
        End If
    End With
End Sub

'卸载飞船上的物资
Private Sub SpacecraftUnloadStorageAll(Spacecraft As Spacecraft)
    Dim i As Long
    
    With Spacecraft
        If .Position.Type = SpacePositionTypeLand Then
            For i = UBound(.Storage) To 1 Step -1
                PlanetAddResource Planets(SpacecraftGetPlanet(Spacecraft)), .Storage(i).Type, .Storage(i).Amont
                SpacecraftDeleteStorage Spacecraft, i
            Next
        End If
            
        Planets(SpacecraftGetPlanet(Spacecraft)).Colonys.Population = Planets(SpacecraftGetPlanet(Spacecraft)).Colonys.Population + .Population
        .Population = 0
    End With
End Sub

'绘制飞船对应的UI
Private Sub SpacecraftDrawUI(Spacecraft As Spacecraft, ByVal X As Long, ByVal Y As Long)
    Dim DrawText As String
    Dim DrawRectangle As Rectangle
    Dim i As Long

    With Spacecraft
        FillColor = vbWhite
        Line (X, Y)-(X + 200, Y + 300), , B
        
        MyPrint .Name, X + 10, Y + 10
        If Not .Enabled Then MyPrint "已禁用", X + 10, CurrentY
        MyPrint "维护费:" & .Maintenance, X + 10, CurrentY
        MyPrint "人口:" & .Population, X + 10, CurrentY
        
        For i = 1 To UBound(.Effects)
            With .Effects(i)
                Select Case .Type
                Case EffectNone
                    MyPrint GetEffectName(.Type), X + 10, CurrentY
                Case EffectResource
                    MyPrint GetResourceName(.EffectResources.Type) & ":" & .EffectResources.Amont, X + 10, CurrentY
                Case Else
                    MyPrint GetEffectName(.Type) & ":" & .Amont, X + 10, CurrentY
                End Select
            End With
        Next
        
        MyPrint "仓储:", X + 10, CurrentY
        
        If UBound(.Storage) = 0 Then
            MyPrint "无", X + 10, CurrentY
        Else
            For i = 1 To UBound(.Storage)
                MyPrint ResourceToString(.Storage(i)), X + 10, CurrentY
            Next
        End If
        
        DrawButtonWithUI RectangleCreate(X + 10, Y + 220, 60, 30), "拆除", "system(""dismantle_spacecraft " & SelectSpacecraft & """)"
        
        If .Construction = 1 Then
            '绘制启/禁用按钮
            If .Enabled Then
                DrawText = "禁用"
            Else
                DrawText = "启用"
            End If
            DrawButtonWithUI RectangleCreate(X + 80, Y + 220, 60, 30), DrawText, "system(""switch_spacecraft_enabled " & SelectSpacecraft & """)"
            
            DrawButtonWithUI RectangleCreate(X + 80, Y + 260, 60, 30), "卸载全部", "system(""unload_spacecraft_storage " & SelectSpacecraft & """)"
            
            DrawButtonWithUI RectangleCreate(X + 80, Y + 300, 60, 30), "打开仓库", "system(""open_spacecraft_storage " & SelectSpacecraft & """)"
            
            DrawRectangle = RectangleCreate(X + 10, Y + 260, 60, 30)
            For i = 1 To UBound(Planets)
                If Not SpacecraftGetPlanet(Spacecraft) = i Then
                    DrawButtonWithUI DrawRectangle, "移动到" & Planets(i).Name, "system(""move_spacecraft_to " & SelectSpacecraft & " " & i & """)"   '绘制移动按钮
                    DrawRectangle.Top = DrawRectangle.Top + 40
                End If
            Next
        End If
    End With
    
    DrawCloseButton X + 170, Y + 5, "clear_select"
End Sub

'创建在太空中的位置
Private Function SpacePositionCreate(PositionType As SpacePositionTypeEnum, ByVal Position1 As Double, Optional Position2 As Double = 0, Optional ByVal Progress As Double = 1) As SpacePosition
    With SpacePositionCreate
        .Type = PositionType
        .Position1 = Position1
        .Position2 = Position2
        .Progress = Progress
    End With
End Function

'获取飞船所在的星球
Private Function SpacePositionGetPlanet(SpacePosition As SpacePosition) As Long
    With SpacePosition
        Select Case .Type
        Case SpacePositionTypeLand
            SpacePositionGetPlanet = .Position1
        Case SpacePositionTypeNavigation
            SpacePositionGetPlanet = 0
        Case SpacePositionTypeSurround
            SpacePositionGetPlanet = .Position1
        End Select
    End With
End Function

Private Sub StarDraw(Star As Star, X As Long, Y As Long)
    With Star
        FillColor = .Color
        Circle (X, Y), 15
        CurrentX = X - 0.5 * TextWidth(.Name)
        CurrentY = Y + 20
        Print .Name
    End With
End Sub

'计算恒星相对太阳的光度
Private Function StarGetRelativeLuminosity(Star As Star) As Double
    StarGetRelativeLuminosity = 10 ^ (Star.Magnitude - 4.83)
End Function

Private Sub SystemDraw(System As System)
    Dim i As Long
    Dim ID As Long
    
    With System
        For i = 1 To UBound(.Stars)
            StarDraw Stars(FindStarWithTag(.Stars(1))), GetCenterX, GetCenterY
        Next
        
        For i = 1 To UBound(.Planets)
            FillStyle = 1
            Circle (GetCenterX, GetCenterY), 40 * i
            FillStyle = 0
            ID = FindPlanetWithTag(.Planets(i))
            PlanetDraw Planets(ID), GetCenterX + 40 * i * Cos(Planets(ID).OrbitRotation), GetCenterY + 40 * i * Sin(Planets(ID).OrbitRotation)
            AddUIObjectList RectangleCreate(GetCenterX - 5 + 40 * i * Cos(Planets(ID).OrbitRotation), GetCenterY - 5 + 40 * i * Sin(Planets(ID).OrbitRotation), 10, 10), "system(""select_planet " & i & """)"
        Next
    End With
End Sub

Private Function Tanh(X As Double) As Double
    Tanh = (1 - Exp(-2 * X)) / (1 + Exp(-2 * X))
End Function

Private Sub UICalculate()
    DrawModuleOffset = DrawModuleOffset * 0.4 + DrawModuleOffsetTarget * 0.6
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 192
        IsShowDebugWindow = Not IsShowDebugWindow
    End Select

    Select Case FormOn
    Case FormGame
        Select Case KeyCode
        Case vbKeySpace
            GameSpeed = 1 - GameSpeed
        Case vbKeyEscape
            IsShowPauseMenu = Not IsShowPauseMenu
        End Select
    End Select
End Sub

Private Sub Form_Load()
    Dim Freq As Currency '每秒计时器次数
    Randomize
    Move Screen.Width * 0.2, Screen.Height * 0.2, Screen.Width * 0.6, Screen.Height * 0.6
    QueryPerformanceFrequency Freq
    FrequencyPerMillisecond = Freq / 1000
    FormOn = FormMainMenu
    
'    DoActionText "msgbox 我是猪"

'    Interpreter.Run "a=4*(6+6)/-3;{b=a*2-a;c=a+b}"
    Interpreter.Run "1/0"
'    Interpreter.Run "prog a:a=3;a=7"
'    Interpreter.Run "{"
'    Interpreter.Run "system(""msgbox 我是猪"")"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Button = Button
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim MouseX As Long
    Dim MouseY As Long
    MouseX = MousePosition.X
    MouseY = MousePosition.Y
    For i = UBound(UIObjectList) To 1 Step -1
        With UIObjectList(i)
            If RectangleIsIn(.RealPosition, MouseX, MouseY) And (Button = .Button Or .Button = 0) Then
                Interpreter.Run .ClickAction
                Exit Sub
            End If
        End With
    Next i
    Button = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Timer1_Timer()
    Dim NewTime As Currency, TickTime As Currency, DrawTime As Currency
    Timer1.Enabled = False
    Do While FormOn <> FormClosed
        QueryPerformanceCounter NewTime
        If (NewTime - DrawTime) / FrequencyPerMillisecond > 30 Then
            DrawTime = NewTime
            Cls
            ReDim UIObjectList(0)
            '绘制窗口到屏幕
            Select Case FormOn
            Case FormMainMenu
                DrawMainMenu
            Case FormStartGame
                DrawStartGame
            Case FormGame
                GetFPS
                If Not IsShowPauseMenu Then
                    If GameSpeed > 0 Then
                        If (NewTime - TickTime) / FrequencyPerMillisecond > 1000 / GameSpeed Then
                            TickTime = NewTime
                            DailyCaculate
                        End If
                    End If
                End If
                DrawGame
            Case FormSettings
                DrawSettings
            End Select
            
            UICalculate
            
            If IsShowDebugWindow Then DrawDebugWindow
        Else
            Sleep 30
        End If
        DoEvents
    Loop
    End
End Sub
