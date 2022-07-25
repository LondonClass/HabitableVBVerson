VERSION 5.00
Begin VB.Form FormMain 
   AutoRedraw      =   -1  'True
   Caption         =   "�˾�"
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8565
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   373
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   571
   StartUpPosition =   3  '����ȱʡ
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

'==ö����������==

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

'��̫���е�λ������
Private Enum SpacePositionTypeEnum
    SpacePositionTypeLand '��½
    SpacePositionTypeSurround '����
    SpacePositionTypeNavigation '����
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

'==�Զ�����������==

Private Type PointApi
    X As Long
    Y As Long
End Type

Private Type Rectangle 'UI����
    Left As Long '�������Ե��������Ե�ľ���
    Top As Long '�����ϱ�Ե��������Ե�ľ���
    Width As Long '���ο��,�����ұ�Ե����������Ե������
    Height As Long '���θ߶�,�����±�Ե��������ϱ�Ե������
End Type

Private Type UIObject 'UI���󣬰��������͡�λ�����С����������Ӧ��ʽ����Ϣ
'    UIType As String 'UI���������,���������UI�����ǰ�ť��ͼ��������ֻ�������������������Ʒ�ʽ
'    'UIType��ֵ��"button"����ť,"clicker"��͸����ť��
'    Info As String '������UI������
    Position As Rectangle 'UI����Ĵ�С��λ��
    RealPosition As Rectangle 'UI����Ĵ�С��λ��
    Parent As Long 'UI�ĸ�����ID
    Button As Long 'UI���������Ӧ����갴��
    ClickAction As String '���UI������ִ�е�ָ��
    Tooltip As String 'UI�������ͣ��ʾ
End Type

Private Type Resource
    Type As ResourceEnum
    Amont As Double
End Type

Private Type Material
    Type As String '����
    Mass As Double
End Type

Private Type MaterialType
    Name As String '����
    State As MaterialStateEnum '��̬
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

'��������
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

'����
Private Type Module
    Type As Long '����
    Size As Long '�����ȼ�
    Construction As Double '�������
    Enabled As Boolean '�Ƿ�����
    EfficiencyModifier As Double 'Ч�ʼӳ�
    Efficiency As Double 'ʵ��Ч��
    Owner As Long '������
    Storage() As Resource '�ִ�
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

'����
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
    RotationPeriod As Double '��ת����
    OrbitRadius As Double '��ת�뾶
    Color As OLE_COLOR '���ǵ���ɫ
    
    OrbitRotation As Double
    
    Colonys As Colony
    Resources() As Double
    Transport() As Double
    Modules() As Module
    Materials() As Material
    
    Market As Market
    
    UtilizingBlock As Long
    BioMass As Double
    HomeWorld As Boolean '�Ƿ�ĸ��
    
    '�����л���
    Housing As Double 'ס������
    Storage As Double '�ܴ洢�ռ�
End Type

'����ϵ
Private Type System
    Name As String
    ID As Long
    Stars() As String
    Planets() As String
End Type

'��̫���е�λ��
Private Type SpacePosition
    Type As SpacePositionTypeEnum 'λ������
    Position1 As Double '���������
    Position2 As Double '�������յ�
    Progress As Double '�������������ľ���0-1
End Type

'����ɴ�
Private Type Spacecraft
    Name As String '����
    Maintenance As Double 'ά����
    Power As Double '����
    Position As SpacePosition 'λ��
    Population As Long '��Ա����
    Effects() As Effect 'Ч��
    Owner As String '������
    Space As Double '�洢�ռ�
    Storage() As Resource '�ִ�����
    Construction As Double '�������
    Enabled As Boolean
End Type

Private Type Technology
    Name As String
    NeedPoints As Long
    IsResearched As Boolean
End Type

Private Type FinancialInfo
    Funds As Double '����
    Contribution As Double '���
    Income As Double '������
    ColonyMaintence As Double 'ֳ���ά��
    Salary As Double '����
    Transport As Double '����
    Expence As Double '��֧��
    NetIncome As Double  '�ܼ�
End Type

Private Type GameEvent
    Title As String
    Content As String
    Options() As String
End Type

'==��������==

Dim FormOn As FormEnum '��ǰ��ʾ����
Dim FrequencyPerMillisecond As Double
Dim Fps As Long

'��Ϸ���ݱ���
Dim GameDate As Date '��Ϸ����
Dim PreviousDate As Date '��Ϸ���ڵ���һ��
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

'UI��ر���
Dim ShowEvents() As Long
Dim ShowingUI As ShowUI
Dim SelectMenuButton As Long '��ʱ��������¼UI����Ĳ˵���
Dim SelectModule As Long
Dim SelectSpacecraft As Long
Dim SelectPlanet As Long
Dim DrawModuleOffsetTarget As Long
Dim DrawModuleOffset As Long
Dim IsShowPauseMenu As Boolean
Dim IsShowDebugWindow As Boolean
Dim IsWriteLog As Boolean

Dim GameSpeed As Long '��Ϸ�ٶ�

Dim GameLog As String '�������ɵ���־�ļ�

'Dim MouseButton As Long
Dim UIObjectList() As UIObject '����UI��������Ӧ���������б�

'==api����==

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long '��ʱ��
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long '��ȡ��ʱ��Ƶ��
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As PointApi) As Long '��ȡ���λ��

'��UI�б������UI������������UI�б��е����
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

'���ѡ��
Private Sub ClearSelect()
    SelectModule = 0
    SelectSpacecraft = 0
End Sub

Private Function CloseEvent(n As Long)
    ReDim ShowEvents(0)
End Function

'����ֳ�������
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

'ÿ�ռ���
Private Sub DailyCaculate()
    Dim i As Long, j As Long
    
    PreviousDate = GameDate
    GameDate = GameDate + 1
    
    '���Ǽ���
    For i = 1 To UBound(Planets)
        DailyCaculatePlanet Planets(i)
    Next
    
    '�¶ȼ���
    If Month(PreviousDate) <> Month(GameDate) Then
        '���������Ϣ
        With FinancialInfo
            Money = Money + Funds
            .Funds = Funds
            Money = Money + Int(10 * Prestige ^ 0.5)
            .Contribution = Int(10 * Prestige ^ 0.5)
            .Income = .Funds + .Contribution
            For i = 1 To UBound(Planets)
                With Planets(i)
                    If Not .HomeWorld Then
                        'ģ�龭�ü���
                        For j = 1 To UBound(.Modules)
                            Money = Money + ModuleTypes(.Modules(j).Type).Maintenance * .Modules(j).Efficiency * .Modules(j).Size
                            FinancialInfo.ColonyMaintence = FinancialInfo.ColonyMaintence + ModuleTypes(.Modules(j).Type).Maintenance * .Modules(j).Efficiency * .Modules(j).Size
                        Next
                        '���ʼ���
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

'�г�ÿ�ռ���
Private Sub DailyCaculateMarket(Market As Market, Planet As Planet)
    Dim Demand As Double
    Dim ActualBought As Double
    Dim TempMoney As Double
    Dim TempGoods As Double
    Dim TempPrice As Double
    
    With Market
        .Money(ResourceFood) = .Money(ResourceFood) + 1 '��ʱ�ṩ������
        
        If .Storage(ResourceFood) > 1 And Planet.Colonys.PopulationMoney > 1 Then
            '���㽻��
            .Prices(ResourceFood) = .Money(ResourceFood) / .Storage(ResourceFood)
            
            '�����˿�����
            Demand = Min(Planet.Colonys.Population, Planet.Colonys.PopulationMoney / .Prices(ResourceFood))
            
            '�ж��г��Ƿ����㹻����
            If Demand < .Storage(ResourceFood) Then
                '�����г�Ӧ�н�Ǯ
                TempMoney = MarketGetProduct(Market, ResourceFood) / (.Storage(ResourceFood) - Demand)
                
                TempPrice = (TempMoney - .Money(ResourceFood)) / Demand
                
                If (TempMoney - .Money(ResourceFood)) < Planet.Colonys.PopulationMoney Then '���Ǯ�Ƿ���
                    ActualBought = Demand
                    
                    Planet.Colonys.PopulationMoney = Planet.Colonys.PopulationMoney - ActualBought * TempPrice
                    .Money(ResourceFood) = .Money(ResourceFood) + ActualBought * TempPrice
                    .Storage(ResourceFood) = .Storage(ResourceFood) - ActualBought
                    Planet.Colonys.PopulationStorage(ResourceFood) = Planet.Colonys.PopulationStorage(ResourceFood) + ActualBought
                Else
                    'ֻ�򲿷֣�������Ǯ������������
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

'ģ��ÿ�ռ���
Private Sub DailyCaculateModule(Module As Module, Planet As Planet, ByVal ID As Long)
    With Module
        '����ģ��
        If .Construction < 1 Then
            .Construction = .Construction + 1 / ModuleTypes(.Type).BuildTime
            If .Construction > 1 Then
                .Construction = 1
                .Size = .Size + 1
            End If
        End If
        
        '���ģ���Ƿ���ȳ�ѹ�����ݻٹ��ȳ�ѹģ��
        If ModuleTypes(.Type).MaxTempreture < Planet.Tempreture Then
            MsgBox Planet.Name & "�ϵĽ���" & ModuleTypes(.Type).Name & "�����¶ȹ��߱���"
            PlanetDeleteMoudle Planet, ID
        End If
        If ModuleTypes(.Type).MaxPressure < PlanetGetPressure(Planet) Then
            MsgBox Planet.Name & "�ϵĽ���" & ModuleTypes(.Type).Name & "������ѹ���󱻻�"
            PlanetDeleteMoudle Planet, ID
        End If
        
        '����ģ��Ч��
        MoudleEfficiencyCalculate Module, Planet, ID

        '����ģ��Ч��
        MoudleEffectCalculate Module, Planet
    End With
End Sub

'����ÿ�ռ���
Private Sub DailyCaculatePlanet(Planet As Planet)
    Dim i As Long
    Dim ResourceSend As Double '��ʱ�������������Դ����
    
    With Planet
        .Housing = 0
        .Storage = 0
        
        '�ɴ�Ч������
        For i = 1 To UBound(Spacecrafts)
            If SpacecraftGetPlanet(Spacecrafts(i)) = PlanetGetID(Planet) Then
                DailyCaculateSpacecraft Spacecrafts(i)
            End If
        Next
    
        '������ת
        .OrbitRotation = .OrbitRotation - 2 * 4 * Atn(1) / PlanetGetOrbitalPeriod(Stars(1), Planet)
        
        '�¶ȼ���
        .Tempreture = .Tempreture - 10 * (.Tempreture / 273.15) ^ 4 * (1 - PlanetGetGreenhouseEffect(Planet)) + 12 * (1 - GetReflectivity(Planet)) * (1.496 / .OrbitRadius) ^ 2
        
        '��̬����
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
        
'        '����г�����
'        If .HomeWorld Then
'            '�����������
'            For i = 1 To UBound(.MarketSupply)
'                .MarketSupply(i) = 0
'            Next
'        End If

        'ģ�����
        For i = 1 To UBound(.Modules)
            DailyCaculateModule .Modules(i), Planet, i
        Next
        If .Tempreture < 0 Then .Tempreture = 0 '��ֹ�¶�Ϊ��
        
        '��Դ�����
        For i = 1 To UBound(.Resources)
            If .Resources(i) > Planet.Storage Then
                .Resources(i) = .Resources(i) - 0.01 * (.Resources(i) - Planet.Storage)
            End If
        Next
    
        '����������Դ
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
        
        '�г�����
        If .HomeWorld Then
            DailyCaculateMarket .Market, Planet
        End If
        
        '�������
'        If Not .HomeWorld Then
'            ResourceSend = Min(.Resources(ResourceFood), .Colonys.Population)
'            .Resources(ResourceFood) = .Resources(ResourceFood) - ResourceSend
'            .Colonys.PopulationStorage(ResourceFood) = .Colonys.PopulationStorage(ResourceFood) + ResourceSend
'
'            ResourceSend = Min(.Resources(ResourceCleanWater), .Colonys.Population)
'            .Resources(ResourceCleanWater) = .Resources(ResourceCleanWater) - ResourceSend
'            .Colonys.PopulationStorage(ResourceCleanWater) = .Colonys.PopulationStorage(ResourceCleanWater) + ResourceSend
'        End If
    
        '�˿ڼ���
        DailyCaculatePopulation .Colonys, Planet
    End With
End Sub

'�˿�ÿ�ռ���
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

'�ɴ�ÿ�ռ���
Private Sub DailyCaculateSpacecraft(Spacecraft As Spacecraft)
    Dim i As Long
    
    With Spacecraft
        '����ɴ�
        If .Construction < 1 Then
            .Construction = .Construction + 1 / 60
            If .Construction > 1 Then
                .Construction = 1
            End If
        End If

        '����ɴ�Ч��
        If SpacecraftGetPlanet(Spacecraft) <> 0 Then
            For i = 1 To UBound(Spacecraft.Effects)
                EffectCalculate Spacecraft.Effects(i), Planets(SpacecraftGetPlanet(Spacecraft)), 1
            Next
        End If
    End With
End Sub

'ɾ��������
Private Sub DeleteSpacecraft(ID As Long)
    Dim i As Long
    For i = ID To UBound(Spacecrafts) - 1
        Spacecrafts(i) = Spacecrafts(i + 1)
    Next
    ReDim Preserve Spacecrafts(UBound(Spacecrafts) - 1)
End Sub

Public Sub DoActionSentence(ByVal Action As String) '����ָ�����,��ִ��
    Dim Words() As String
    Dim Name As String
    Dim Parameters() As Variant
    Dim i As Long
    
    '��ָ���б��зָÿһ�ж������������ٴ�ӡ
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

Private Sub DoSystemAction(ByVal Name As String, Parameters() As Variant) '�������Ӧ
    Dim ID As Long
    
    Select Case Name
    Case "" '��Ч��
    Case "none" '��Ч��
    Case "develop_population"
'    Case 1 '��չ�˿� ����0Ϊ������ ����1Ϊ�˿�����
        If Money > 10 Then
            ID = FindPlanetWithTag(CStr(Parameters(0)))
            With Planets(ID)
                If .Colonys.Population < Planets(ID).Housing Then
                    Money = Money - Parameters(1)
                    .Colonys.Population = .Colonys.Population + Parameters(1)
                Else
                    MsgBox "��ס�ռ䲻��"
                End If
            End With
        Else
            MsgBox "�ʽ���"
        End If
    Case "bulid_module"
'    Case 2 '���콨�� ����0Ϊ������ ����1Ϊ��������
        PlanetBuildModule Planets(FindPlanetWithTag(CStr(Parameters(0)))), Parameters(1)
    Case "change_showing_ui"
'    Case 3 '�л�������ʾ��UI ����0Ϊ�л�����UI���
        ShowingUI = Parameters(0)
        SelectMenuButton = 0
    Case "set_select_module"
'    Case 4 'ѡ���� ����0Ϊ�������
        SelectModule = Parameters(0)
    Case "switch_select_module" '�л�ѡ�еĽ���
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
    Case "switch_select_spacecraft" '�л�ѡ�еķɴ�
        If SelectSpacecraft <> Parameters(0) Then
            ClearSelect
            SelectSpacecraft = Parameters(0)
        Else
            SelectSpacecraft = 0
        End If
'    Case 5 '������ҳ ����0Ϊ��ҳ����
    Case "change_module_offset"
        DrawModuleOffsetTarget = DrawModuleOffsetTarget + Parameters(0)
        If DrawModuleOffsetTarget < 0 Then DrawModuleOffsetTarget = 0
    Case "dismantle_module"
'    Case 6 '������� ����0Ϊ������ ����1Ϊ�������
        PlanetDeleteMoudle Planets(FindPlanetWithTag(CStr(Parameters(0)))), CLng(Parameters(1))
    Case "dismantle_spacecraft"
        DeleteSpacecraft CLng(Parameters(0))
    Case "transport_resources"
'    Case 7 'һ����������Դ ����0Ϊ������ ����1Ϊ��Դ��� ����2Ϊ����
        If Money > 10 Then
            With Planets(FindPlanetWithTag(CStr(Parameters(0))))
                If UBound(.Resources) >= Parameters(1) Then
                    Money = Money - 10
                    .Resources(Parameters(1)) = .Resources(Parameters(1)) + Parameters(2)
                Else
                    MsgBox "�±�Խ��"
                End If
            End With
        Else
            MsgBox "�ʽ���"
        End If
    Case "add_transport_resources"
'    Case 8 '���ӳ���������Դ ����0Ϊ������ ����1Ϊ��Դ��� ����2Ϊ����
'    Case 9 '���ٳ���������Դ
        With Planets(FindPlanetWithTag(CStr(Parameters(0))))
            If UBound(.Resources) >= Parameters(1) Then
                .Transport(Parameters(1)) = .Transport(Parameters(1)) + Parameters(2)
            Else
                MsgBox "������Դ�����±�Խ��"
            End If
        End With
    Case "change_window"
'        Case 10 '�л����� ����0Ϊ�л����Ĵ��ڱ��
        FormOn = Parameters(0)
        If Parameters(0) = FormGame Then
            GameInitialization
        End If
    Case "add_utilizing_block"
'    Case 11 '���ӽ����ռ� ����0Ϊ��ӽ����ռ��������
        ID = FindPlanetWithTag(CStr(Parameters(0)))
        With Planets(ID)
            If Money > GetBlockCost(Planets(ID)) Then
                Money = Money - GetBlockCost(Planets(ID))
                .UtilizingBlock = .UtilizingBlock + 1
            Else
                MsgBox "�ʽ���"
            End If
        End With
    Case "select_planet"
'    Case 12 'ѡ������
        SelectPlanet = Parameters(0)
        ShowingUI = ShowUINone
    Case "close_event"
'    Case 13 '�ر��¼� ����0Ϊʱ����
        CloseEvent CLng(Parameters(0))
'    Case 14 '�ı���Ϸ�ٶ�
    Case "set_speed"
        GameSpeed = Parameters(0)
    Case "swith_pause_menu"
'    Case 15 '��ʾ/������ͣ����
        IsShowPauseMenu = Not IsShowPauseMenu
    Case "switch_module_enabled"
'    Case 16 '��/���ý��� ����0Ϊ������ ����1Ϊ�������
        ID = FindPlanetWithTag(CStr(Parameters(0)))
        Planets(ID).Modules(Parameters(1)).Enabled = Not Planets(ID).Modules(Parameters(1)).Enabled
        
    Case "switch_spacecraft_enabled" '��/���÷ɴ� ����0Ϊ�ɴ����
        Spacecrafts(Parameters(0)).Enabled = Not Spacecrafts(Parameters(0)).Enabled
    
    Case "move_spacecraft_to" '��/���÷ɴ� ����0Ϊ�ɴ���� ����1Ϊ������
        Spacecrafts(Parameters(0)).Position.Position1 = Parameters(1)
    
    Case "unload_spacecraft_storage" 'ж�طɴ��ϵ����� ����0Ϊ�ɴ����
        SpacecraftUnloadStorageAll Spacecrafts(Parameters(0))
    
    Case "expand_module"
'    Case 17 '�������� ����0Ϊ������ ����1Ϊ�������
        ExpansionModule Planets(FindPlanetWithTag(CStr(Parameters(0)))), Parameters(1)
    Case "set_menu_button"
'    Case 18 '�ı���水ť
        SelectMenuButton = Parameters(0)
'    Case 19 '����ֿ���Դ
    Case "save_game" '������Ϸ
        SaveGame
    Case "msgbox" '������Ϸ
        MsgBox Parameters(0)
    End Select
End Sub

'���Ƶײ��˵�
Private Sub DrawBottomBar(Position As Rectangle)
    With Position
        DrawButtonWithUI Position, "����", "system(""change_showing_ui " & ShowUIFinance & """)"
        .Left = .Left + .Width + 10
        
        DrawButtonWithUI Position, "�˿�", "system(""change_showing_ui " & ShowUIPopulation & """)"
        .Left = .Left + .Width + 10
        
        DrawButtonWithUI Position, "�о�", "system(""change_showing_ui " & ShowUIResearch & """)"
        .Left = .Left + .Width + 10
        
        If SelectPlanet > 0 Then
            DrawButtonWithUI Position, "����", "system(""change_showing_ui " & ShowUIPlanet & """)"
            .Left = .Left + .Width + 10
            
            DrawButtonWithUI Position, "��Դ", "system(""change_showing_ui " & ShowUIResource & """)"
        End If
    End With
End Sub

'���Ʊ�׼��ť
Private Sub DrawButton(Position As Rectangle, ByVal Text As String, Optional ByVal Color As OLE_COLOR = &HC0C0C0)
    RectangleDraw Position, Color '����
    MyPrint Text, Position.Left + 0.5 * Position.Width, Position.Top + 0.5 * Position.Height, 1 '��������
End Sub

'���Ʊ�׼��ť�����UI
Private Sub DrawButtonWithUI(Position As Rectangle, Optional ByVal Text As String = "NO_TEXT", Optional ByVal ClickAction As String = "", Optional ByVal Tooltip As String = "", Optional ByVal Color As OLE_COLOR = &HC0C0C0, Optional ByVal Parent As Long = 0, Optional ByVal Button As Long = 1)
    DrawButton Position, Text, Color
    AddUIObjectList Position, ClickAction, Button, Tooltip, Parent
End Sub

Private Sub DrawCloseButton(ByVal X As Long, ByVal Y As Long, Optional ByVal ClickAction As String = "system(""change_showing_ui " & ShowUINone & """)", Optional ByVal Size As Long = 15) '���ƹرհ�ť
    DrawButtonWithUI RectangleCreate(X, Y, Size, Size), "��", ClickAction, , RGB(255, 64, 64)
End Sub

Private Sub DrawDebugWindow() '����debug����
    Dim ButtonPosition As Rectangle
    
    FillColor = vbWhite
    Line (50, 50)-(300, GetCenterY + 100), , B
End Sub

Private Sub DrawEvent(ByVal X As Long, ByVal Y As Long, ByVal ID As Long) '�����¼�����
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

Private Sub DrawGame() '������Ϸ���浽��Ļ
    '������Ϸ���岿��
    If SelectPlanet = 0 Then
        SystemDraw System
    Else
        DrawModules Planets(SelectPlanet)
        DrawSpacecrafts Planets(SelectPlanet)
    End If
    '������ϷUI
    DrawGameUI
End Sub

Private Sub DrawGameUI()
    Dim i As Long
    Dim ButtonPosition As Rectangle
    
    CurrentX = 0
    CurrentY = 0
    Print "FPS:" & Fps
    Print "����:" & GameDate
    Print "�ʽ�:" & Int(Money)
    Print "����:" & Int(Prestige)
    Print "�о�����:" & Int(ResearchPoint)
    
    '��ʾѡ��ģ��Ч��
    If SelectPlanet > 0 Then
        If SelectModule > 0 And SelectModule <= UBound(Planets(SelectPlanet).Modules) Then
            ModuleDrawUI Planets(SelectPlanet).Modules(SelectModule), 500, 10
        End If
    End If
    
    '��ʾѡ�к�����Ч��
    If SelectSpacecraft > 0 And SelectSpacecraft <= UBound(Spacecrafts) Then
        SpacecraftDrawUI Spacecrafts(SelectSpacecraft), 500, 10
    End If
    
    DrawSpeedControlBar RectangleCreate(ScaleWidth - 170, 20, 30, 20) '�����ٶȿ�������λ����Ļ���Ͻ�
    
    '���ƴ򿪵�UI����
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
    
    DrawBottomBar RectangleCreate(30, ScaleHeight - 70, 70, 40) '���Ƶײ��˵�
    
    If SelectPlanet > 0 Then DrawButtonWithUI RectangleCreate(ScaleWidth - 50, 50, 30, 30), "��ϵ", "system(""select_planet 0"")" '���Ʒ�����ϵ��ť
    
    DrawToolTips '������ʾ
    
    '��ʾ�¼�
    For i = 1 To UBound(ShowEvents)
        DrawEvent GetCenterX - 150, GetCenterY - 100, ShowEvents(i)
    Next
     
    If IsShowPauseMenu Then DrawPauseMenu
End Sub

Private Sub DrawGameUIFinance(ByVal X As Long, ByVal Y As Long)
    Dim DrawPosition As Rectangle
    
    FillColor = vbWhite
    Line (X, Y)-(X + 400, Y + 300), , B
    
    MyPrint "����", X + 10, Y + 15
    
    With PreviousFinancialInfo
        CurrentX = X + 10
        Print "����:" & Format(.Funds, "0.00")
        CurrentX = X + 10
        Print "���:" & Format(.Contribution, "0.00")
        CurrentX = X + 10
        Print "������:" & Format(.Income, "0.00")
        CurrentX = X + 10
        Print "ֳ���ά��:" & Format(.ColonyMaintence, "0.00")
        CurrentX = X + 10
        Print "����:" & Format(.Salary, "0.00")
        CurrentX = X + 10
        Print "����:" & Format(.Transport, "0.00")
        CurrentX = X + 10
        Print "��֧��:" & Format(.Expence, "0.00")
        CurrentX = X + 10
        Print "�ܼ�:" & Format(.NetIncome, "0.00")
    End With
    
    DrawCloseButton X + 330, Y + 5
End Sub

Private Sub DrawUILoadGame(ByVal X As Long, ByVal Y As Long)
    FillColor = vbWhite
    Line (X, Y)-(X + 400, Y + 300), , B
    
    MyPrint "û�д浵", X + 10, Y + 15
End Sub

Private Sub DrawUINewCampaign(ByVal X As Long, ByVal Y As Long)
    FillColor = vbWhite
    Line (X, Y)-(X + 400, Y + 300), , B
    
    CurrentX = X + 10
    CurrentY = Y + 15
    Print "û�г���"
End Sub

Private Sub DrawUINewGame(ByVal X As Long, ByVal Y As Long)
    FillColor = vbWhite
    Line (X, Y)-(X + 400, Y + 300), , B
    
    CurrentX = X + 10
    CurrentY = Y + 15
    Print "����Ϸ"
End Sub

Private Sub DrawGameUIPlanet(Planet As Planet, ByVal X As Long, ByVal Y As Long)
    Dim DrawPosition As Rectangle
    Dim i As Long
    
    FillColor = vbWhite
    Line (X, Y)-(X + 400, Y + 300), , B
    
    CurrentX = X + 10
    CurrentY = Y + 15
    Print "����"
    
    If SelectPlanet = 0 Then Exit Sub
    
    Select Case SelectMenuButton
    Case 0
        DrawButtonWithUI RectangleCreate(X + 10, Y + 40, 60, 20), "������Ϣ", "system(""set_menu_button 0"")", , RGB(128, 128, 128)
        DrawButtonWithUI RectangleCreate(X + 80, Y + 40, 40, 20), "����", "system(""set_menu_button 1"")"
        DrawButtonWithUI RectangleCreate(X + 130, Y + 40, 40, 20), "�г�", "system(""set_menu_button 2"")"
    
        With Planet
            CurrentX = X + 10
            CurrentY = Y + 70
            Print .Name
            
            CurrentX = X + 10
            Print "��̬���:"
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
                Print "��"
            End If
            
            CurrentX = X + 10
            Print "�������:"
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
                Print "��"
            End If
            
            CurrentX = X + 10
            Print "�������:"
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
                Print "��"
            End If
            
            CurrentX = X + 10
            Print "����:" & Format(PlanetGetOxygen(Planet), "0.000000%");
            CurrentX = X + 140
            If PlanetGetOxygen(Planet) < 0.1 Then
                Print "��������"
            Else
                If PlanetGetOxygen(Planet) <= 0.32 Then
                    If PlanetGetOxygen(Planet) >= 0.18 And PlanetGetOxygen(Planet) <= 0.24 Then
                        Print "�������˾�ס"
                    Else
                        Print "��������ֲ������"
                    End If
                Else
                    Print "��������"
                End If
            End If
            CurrentX = X + 10
            Print "ˮ:" & Format(PlanetGetWater(Planet), "0.000000%");
            CurrentX = X + 140
            If PlanetGetWater(Planet) < 0.1 Then
                Print "ˮ�ֹ���"
            Else
                If PlanetGetWater(Planet) <= 0.9 Then
                    If PlanetGetWater(Planet) >= 0.25 And PlanetGetWater(Planet) <= 0.75 Then
                        Print "ˮ�����˾�ס"
                    Else
                        Print "ˮ������ֲ������"
                    End If
                Else
                    Print "ˮ�ֹ���"
                End If
            End If
            CurrentX = X + 10
            Print "������" & Format(GetReflectivity(Planet), "0.00%")
            CurrentX = X + 10
            Print "�¶�:" & Format(.Tempreture - 273.15, "0.00") & "��";
            CurrentX = X + 140
            If Planet.Tempreture < 200 Then
                Print "�¶ȹ���"
            Else
                If Planet.Tempreture <= 374 Then
                    If Planet.Tempreture >= 237 And Planet.Tempreture <= 337 Then
                        Print "�¶����˾�ס"
                    Else
                        Print "�¶�����ֲ������"
                    End If
                Else
                    Print "�¶ȹ���"
                End If
            End If
            CurrentX = X + 10
            Print "��ѹ:" & Format(PlanetGetPressure(Planet) / 1000, "0.00") & "kPa";
            CurrentX = X + 140
            If PlanetGetPressure(Planet) < 10000 Then
                Print "��ѹ����"
            Else
                If PlanetGetPressure(Planet) <= 190000 Then
                    If PlanetGetPressure(Planet) >= 50000 And PlanetGetPressure(Planet) <= 150000 Then
                        Print "��ѹ���˾�ס"
                    Else
                        Print "��ѹ����ֲ������"
                    End If
                Else
                    Print "��ѹ����"
                End If
            End If
            CurrentX = X + 10
            Print "����ЧӦ:" & Format(PlanetGetGreenhouseEffect(Planet), "0.00%")
            CurrentX = X + 10
            Print "��ת����:" & Format(PlanetGetOrbitalPeriod(Stars(1), Planet), "0.00") & "��"
            CurrentX = X + 10
            Print "��ת����:" & Format(.RotationPeriod, "0.00") & "��"
            CurrentX = X + 10
            Print "̫����:" & Format(PlanetGetSolarPower(Stars(1), Planet), "0.00%")
        End With
        
        CurrentX = X + 10
        Print "��������:" & GetPowerProduce(Planet)
        CurrentX = X + 10
        Print "��������:" & GetPowerUse(Planet)
        CurrentX = X + 10
        Print "����������:" & Format(GetPowerAdyquacy(Planet), "0.00%")
        
        CurrentX = X + 10
        Print "����:" & Format(GetGravity(Planet), "0.00") & "m/s^2"
        
        CurrentX = X + 10
        Print "����:" & PlanetGetUsedBlock(Planet) & "/" & Planet.UtilizingBlock
        
        CurrentX = X + 10
        Print "������:" & NumberFormat(PlanetGetEvaporation(Planet))
        
        CurrentX = X + 10
        Print "��ˮ�ܶ�:" & NumberFormat(PlanetGetRainfallDensity(Planet))
        
        CurrentX = X + 10
        Print "������:" & NumberFormat(PlanetGetRunoff(Planet))
        
        DrawButtonWithUI RectangleCreate(X + 240, Y + 40, 80, 40), "���ӽ����ռ�" & vbCrLf & "����:" & GetBlockCost(Planet), "add_utilizing_block " & SelectPlanet
    Case 1
        DrawButtonWithUI RectangleCreate(X + 10, Y + 40, 60, 20), "������Ϣ", "system(""set_menu_button 0"")"
        DrawButtonWithUI RectangleCreate(X + 80, Y + 40, 40, 20), "����", "system(""set_menu_button 1"")", , RGB(128, 128, 128)
        DrawButtonWithUI RectangleCreate(X + 130, Y + 40, 40, 20), "�г�", "system(""set_menu_button 2"")"
        
        With Planets(SelectPlanet)
            Print
            CurrentX = X + 10
            Print .Name
            CurrentX = X + 20
            Print "���˿�:" & NumberFormat(.Colonys.Population) & "/" & NumberFormat(Planets(SelectPlanet).Housing)
            CurrentX = X + 20
            If Planets(SelectPlanet).Housing < .Colonys.Population Then
                Print "��ס�ռ䲻��"
            Else
                If Planets(SelectPlanet).Housing = .Colonys.Population Then
                    Print "��ס�ռ�����"
                Else
                    Print "��ס�ռ����"
                End If
            End If
        
            CurrentX = X + 20
            Print "��λ��:" & GetStaffNeed(Planets(SelectPlanet))
            CurrentX = X + 20
            Print "�ڸ���:" & Format(GetStaffAdyquacy(Planets(SelectPlanet)), "0.00%")
            
            CurrentX = X + 20
            Print "��Ȼ����:" & NumberFormat(GetGrowthRate(Planets(SelectPlanet)) * .Colonys.Population)
            
            CurrentX = X + 20
            Print GetResourceName(ResourceFood) & "����:" & NumberFormat(-.Colonys.Population);
            
            If .Resources(ResourceFood) < .Colonys.Population Then
                Print " ȱ��" & GetResourceName(ResourceFood);
                Print " ��Ϊȱ��" & GetResourceName(ResourceFood) & "������:" & NumberFormat(-.Colonys.Population * 0.02 * (1 - .Resources(ResourceFood) / .Colonys.Population))
            Else
                Print " " & GetResourceName(ResourceFood) & "����"
            End If
            
            CurrentX = X + 20
            Print GetResourceName(ResourceCleanWater) & "����:" & NumberFormat(-.Colonys.Population);
            
            If .Resources(ResourceCleanWater) < .Colonys.Population Then
                Print " ȱ��" & GetResourceName(ResourceCleanWater);
                Print " ��Ϊȱ��" & GetResourceName(ResourceCleanWater) & "������:" & NumberFormat(-.Colonys.Population * 0.02 * (1 - .Resources(ResourceCleanWater) / .Colonys.Population))
            Else
                Print " " & GetResourceName(ResourceCleanWater) & "����"
            End If
        End With
    Case 2
        DrawButtonWithUI RectangleCreate(X + 10, Y + 40, 60, 20), "������Ϣ", "system(""set_menu_button 0"")"
        DrawButtonWithUI RectangleCreate(X + 80, Y + 40, 40, 20), "����", "system(""set_menu_button 1"")"
        DrawButtonWithUI RectangleCreate(X + 130, Y + 40, 40, 20), "�г�", "system(""set_menu_button 2"")", , RGB(128, 128, 128)

        For i = 1 To UBound(Planet.Resources)
            FillColor = vbWhite
            Line (X + 10, Y + 20 + 50 * i)-(X + 330, Y + 60 + 50 * i), , B
            
            MyPrint GetResourceName(i), X + 10, Y + 50 * i + 25
            
            CurrentY = Y + 50 * i + 25
            MyPrint "�г��洢:" & NumberFormat(Planet.Market.Storage(i)), X + 80, CurrentY
            MyPrint "�г���Ǯ:" & NumberFormat(Planet.Market.Money(i)), X + 80, CurrentY
            If Planet.Market.Prices(i) < 0.01 And Planet.Market.Prices(i) <> 0 Then
                MyPrint "�۸�:" & Format(Planet.Market.Prices(i), "0.000E-00"), X + 80, CurrentY
            Else
                MyPrint "�۸�:" & Format(Planet.Market.Prices(i), "0.000"), X + 80, CurrentY
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
    
    MyPrint "�˿�", X + 10, Y + 15
    
    If SelectPlanet = 0 Then
        For i = 1 To UBound(Planets)
            With Planets(i)
                Print
                CurrentX = X + 10
                Print .Name
                CurrentX = X + 20
                Print "���˿�:" & NumberFormat(.Colonys.Population) & "/" & NumberFormat(Planets(i).Housing) & " ";
                If Planets(i).Housing < .Colonys.Population Then
                    Print "��ס�ռ䲻��"
                Else
                    If Planets(i).Housing = .Colonys.Population Then
                        Print "��ס�ռ�����"
                    Else
                        Print "��ס�ռ����"
                    End If
                End If
            
                CurrentX = X + 20
                Print "��λ��:" & GetStaffNeed(Planets(i))
                CurrentX = X + 20
                Print "�ڸ���:" & Format(GetStaffAdyquacy(Planets(i)), "0.00%")
                
                CurrentX = X + 20
                Print "��Ȼ����:" & NumberFormat(GetGrowthRate(Planets(i)) * .Colonys.Population)
                
                CurrentX = X + 20
                Print GetResourceName(ResourceFood) & "����:" & NumberFormat(-.Colonys.Population);
                
                If .Resources(ResourceFood) < .Colonys.Population Then
                    Print " ȱ��" & GetResourceName(ResourceFood);
                    Print " ��Ϊȱ��" & GetResourceName(ResourceFood) & "������:" & -Int(.Colonys.Population * 0.05)
                Else
                    Print " " & GetResourceName(ResourceFood) & "����"
                End If
                
                CurrentX = X + 20
                Print GetResourceName(ResourceCleanWater) & "����:" & NumberFormat(-.Colonys.Population);
                
                If .Resources(ResourceCleanWater) < .Colonys.Population Then
                    Print " ȱ��" & GetResourceName(ResourceCleanWater);
                    Print " ��Ϊȱ��" & GetResourceName(ResourceCleanWater) & "������:" & -Int(.Colonys.Population * 0.05)
                Else
                    Print " " & GetResourceName(ResourceCleanWater) & "����"
                End If
            End With
        Next
    Else
        With Planets(SelectPlanet)
            Print
            CurrentX = X + 10
            Print .Name
            CurrentX = X + 20
            Print "���˿�:" & NumberFormat(.Colonys.Population) & "/" & NumberFormat(Planets(SelectPlanet).Housing) & " ";
            If Planets(SelectPlanet).Housing < .Colonys.Population Then
                Print "��ס�ռ䲻��"
            Else
                If Planets(SelectPlanet).Housing = .Colonys.Population Then
                    Print "��ס�ռ�����"
                Else
                    Print "��ס�ռ����"
                End If
            End If
        
            CurrentX = X + 20
            Print "��λ��:" & GetStaffNeed(Planets(SelectPlanet))
            CurrentX = X + 20
            Print "�ڸ���:" & Format(GetStaffAdyquacy(Planets(SelectPlanet)), "0.00%")
            
            CurrentX = X + 20
            Print "��Ȼ����:" & NumberFormat(GetGrowthRate(Planets(SelectPlanet)) * .Colonys.Population)
            
            CurrentX = X + 20
            Print GetResourceName(ResourceFood) & "����:" & NumberFormat(-.Colonys.Population);
            
            If .Resources(ResourceFood) < .Colonys.Population Then
                Print " ȱ��" & GetResourceName(ResourceFood);
                Print " ��Ϊȱ��" & GetResourceName(ResourceFood) & "������:" & NumberFormat(-.Colonys.Population * 0.05)
            Else
                Print " " & GetResourceName(ResourceFood) & "����"
            End If
            
            CurrentX = X + 20
            Print GetResourceName(ResourceCleanWater) & "����:" & NumberFormat(-.Colonys.Population);
            
            If .Resources(ResourceCleanWater) < .Colonys.Population Then
                Print " ȱ��" & GetResourceName(ResourceCleanWater);
                Print " ��Ϊȱ��" & GetResourceName(ResourceCleanWater) & "������:" & NumberFormat(-.Colonys.Population * 0.05)
            Else
                Print " " & GetResourceName(ResourceCleanWater) & "����"
            End If
        
            MyPrint "�ֽ�:" & NumberFormat(.Colonys.PopulationMoney), X + 20, CurrentY
            MyPrint "�Ẕ̌���:" & Format(.Colonys.PopulationReligion, "0.00%"), X + 20, CurrentY
            MyPrint "ƽ�ȶ�:" & Format(.Colonys.Equality, "0.00%"), X + 20, CurrentY
            
            For i = 1 To UBound(.Colonys.PopulationStorage)
                If .Colonys.PopulationStorage(i) > 0 Then
                    MyPrint "�洢" & GetResourceName(i) & ":" & NumberFormat(.Colonys.PopulationStorage(i)), X + 20, CurrentY
                End If
            Next
        End With
        
        DrawButtonWithUI RectangleCreate(X + 160, Y + 40, 60, 40), "��չ�˿�" & vbCrLf & "����:" & 10, "system(""develop_population " & SelectPlanet & " 10"")", "����1�˿�"
    End If
    
    DrawCloseButton X + 330, Y + 5
End Sub


Private Sub DrawUIResearch(ByVal X As Long, ByVal Y As Long)
    Dim DrawPosition As Rectangle
    
    FillColor = vbWhite
    Line (X, Y)-(X + 400, Y + 300), , B
    
    MyPrint "�о�", X + 10, Y + 15
    
    DrawCloseButton X + 330, Y + 5
End Sub

Private Sub DrawGameUIResource(Planet As Planet, ByVal X As Long, ByVal Y As Long)
    Dim i As Long
    Dim Count As Long
    Dim DrawPosition As Rectangle
    Dim ButtonSize As Long '��Ʒͼ���С
    Dim ButtonNum As Long 'ÿ����ʾ����Ʒͼ������
    
    FillColor = vbWhite
    Line (X, Y)-(X + 400, Y + 300), , B
    MyPrint "��Դ", X + 10, Y + 15
    
    Select Case SelectMenuButton
    Case 0
        DrawButtonWithUI RectangleCreate(X + 50, Y + 10, 40, 20), "����", "system(""set_menu_button 1"")"
        
        ButtonSize = 40
        ButtonNum = 7
        For i = 1 To UBound(Planet.Resources)
            '������Ʒͼ��
            DrawPosition = RectangleCreate(X + 10 + (ButtonSize + 10) * ((i - 1) Mod ButtonNum), Y + 40 + (ButtonSize + 30) * ((i - 1) \ ButtonNum), ButtonSize, ButtonSize)
            DrawButtonWithUI DrawPosition, GetResourceName(i), , , RGB(128, 192, 128)
            
            '����Ʒͼ���·�������Ʒ����
            MyPrint NumberFormat(Planet.Resources(i)), DrawPosition.Left + ButtonSize / 2 - TextWidth(NumberFormat(Planet.Resources(i))) / 2, DrawPosition.Top + ButtonSize
        Next
        MyPrint "��Դ����:" & Planet.Storage, X + 10, Y + 280
        
    Case 1
        DrawButtonWithUI RectangleCreate(X + 50, Y + 10, 40, 20), "����", "system(""set_menu_button 0"")", , RGB(128, 128, 128)

        Count = 0
        For i = 1 To UBound(Planet.Resources)
            Count = Count + 1
            FillColor = vbWhite
            Line (X + 10, Y - 10 + 50 * Count)-(X + 330, Y + 30 + 50 * Count), , B
            
            MyPrint GetResourceName(i), X + 10, Y + 50 * Count - 5
            
            MyPrint "��������:" & Planet.Transport(i), X + 80, CurrentY
            
            MyPrint "�ʽ�:" & -Planet.Transport(i), X + 80, CurrentY
            
            DrawButtonWithUI RectangleCreate(X + 180, Y - 5 + 50 * Count, 40, 30), "����", "system(""add_transport_resources " & Planets(SelectPlanet).Tag & " " & i & " 10"")"
            DrawButtonWithUI RectangleCreate(X + 230, Y - 5 + 50 * Count, 40, 30), "����", "system(""add_transport_resources " & Planets(SelectPlanet).Tag & " " & i & " -10"")"
            DrawButtonWithUI RectangleCreate(X + 280, Y - 5 + 50 * Count, 40, 30), "һ����", "system(""transport_resources " & Planets(SelectPlanet).Tag & " " & i & " 10"")"
        Next
        
'        Count = Count + 1
'        FillColor = vbWhite
'
'        DrawPosition = RectangleCreate(X + 10, Y + 50 * Count, 290, 40)
'        DrawButton DrawPosition, "���������"
'        AddUIObjectList DrawPosition, 1, 8, i
    End Select
    
    DrawCloseButton X + 330, Y + 5
End Sub

'���ƽ̳̽���
Private Sub DrawUITutorial(ByVal X As Long, ByVal Y As Long)
    FillColor = vbWhite
    Line (X, Y)-(X + 400, Y + 300), , B
    
    MyPrint "���Ͻ���ʾ��ϸ��Ϣ", X + 10, Y + 15
    MyPrint "�����������������", X + 10, CurrentY
End Sub
    
Private Sub DrawMainMenu() '�������˵�����Ļ
    Dim ButtonPosition As Rectangle

    Font = "΢���ź�"
    ForeColor = vbYellow
    FontSize = 28
    MyPrint "�˾�", GetCenterX, GetCenterY - 110, 1
    
    ForeColor = vbBlack
    FillStyle = 0
    FontSize = 9
    Font = "����"
    ButtonPosition = RectangleCreate(GetCenterX - 90, GetCenterY - 40, 180, 40)
    DrawButtonWithUI ButtonPosition, "��ʼ��Ϸ", "system(""change_window " & FormStartGame & """)"
    
    ButtonPosition = RectangleCreate(GetCenterX - 90, GetCenterY + 20, 180, 40)
    DrawButtonWithUI ButtonPosition, "����", "system(""change_window " & FormSettings & """)"
    
    ButtonPosition = RectangleCreate(GetCenterX - 90, GetCenterY + 80, 180, 40)
    DrawButtonWithUI ButtonPosition, "�˳�", "system(""change_window " & FormClosed & """)"
    
    MyPrint "�˾� Demo20220722a", 0, ScaleHeight - TextHeight("�˾�")
End Sub

Private Sub DrawMenuUI() '���Ʋ˵�UI
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
        '��ҳ��
        DrawButtonWithUI RectangleCreate(.Left + .Width + 10, .Top, .Height, .Height), "�Ϸ�", "system(""change_module_offset " & -300 & """)"
        DrawButtonWithUI RectangleCreate(.Left + .Width + 10, .Top + .Height + 10, .Height, .Height), "�·�", "system(""change_module_offset " & 300 & """)"
        
        DrawPosition = MoudlePosition
        DrawPosition.Top = DrawPosition.Top - DrawModuleOffset
        
        For i = 1 To UBound(Planet.Modules)
            '��ʾģ��ѡ�к�Ļ�ɫ��ʾ��
            If i = SelectModule Then
                RectangleDraw RectangleCreate(DrawPosition.Left - 3, DrawPosition.Top - 3, DrawPosition.Width + 6, DrawPosition.Height + 6), vbYellow
            End If
            
            '��ʾģ��
            If Planet.Modules(i).Construction < 1 Then
                MoudleText = ModuleTypes(Planet.Modules(i).Type).Name & "(������)"
            Else
                If Planet.Modules(i).Enabled = True Then
                    MoudleText = ModuleTypes(Planet.Modules(i).Type).Name
                Else
                    MoudleText = ModuleTypes(Planet.Modules(i).Type).Name & "(�ѽ���)"
                End If
            End If
            DrawButtonWithUI DrawPosition, MoudleText, "system(""switch_select_module " & i & """)"
            DrawPosition.Top = DrawPosition.Top + DrawPosition.Height + 10
        Next
        
        '���ƽ���ģ�鰴ť
        For i = 1 To UBound(ModuleTypes)
            DrawButtonWithUI DrawPosition, "����" & ModuleTypes(i).Name, "system(""bulid_module " & Planets(SelectPlanet).Tag & " " & i & """)"
            DrawPosition.Top = DrawPosition.Top + DrawPosition.Height + 10
        Next
    End With
End Sub

Private Sub DrawPauseMenu() '������ͣ�˵�
    Dim ButtonPosition As Rectangle
    
    FillColor = vbWhite
    Line (GetCenterX - 150, GetCenterY - 100)-(GetCenterX + 150, GetCenterY + 100), , B
    
    DrawButtonWithUI RectangleCreate(GetCenterX - 90, GetCenterY - 80, 180, 40), "������Ϸ", "system(""swith_pause_menu" & """)"
    
    DrawButtonWithUI RectangleCreate(GetCenterX - 90, GetCenterY - 20, 180, 40), "������Ϸ", "system(""save_game" & """)"
    
'    ButtonPosition = RectangleCreate(GetCenterX - 90, GetCenterY - 20, 180, 40)
'    DrawButton ButtonPosition, "����"
'    AddUIObjectList ButtonPosition, 1, 10, FormSettings

    DrawButtonWithUI RectangleCreate(GetCenterX - 90, GetCenterY + 40, 180, 40), "�������˵�", "change_window " & FormMainMenu
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

Private Sub DrawSettings() '�������ý���
    Dim Position As Rectangle
    
    ReDim UIObjectList(0)
    CurrentX = 80
    CurrentY = 40
    Print "��������"
    Position = RectangleCreate(GetCenterX - 180, ScaleHeight - 80, 360, 50)
    DrawButton Position, "���"
    AddUIObjectList Position, "system(""change_window " & FormMainMenu & """)"
End Sub

'���ƺ���������
Private Sub DrawSpacecrafts(Planet As Planet)
    Dim i As Long
    Dim MoudlePosition As Rectangle
    Dim DrawPosition As Rectangle
    
    MoudlePosition = RectangleCreate(300, 10, 100, 30)
    
    With MoudlePosition
        DrawPosition = MoudlePosition
        
        For i = 1 To UBound(Spacecrafts)
            If SpacecraftGetPlanet(Spacecrafts(i)) = PlanetGetID(Planet) Then
                '��ʾ�ɴ�ѡ�к�Ļ�ɫ��ʾ��
                If i = SelectSpacecraft Then
                    RectangleDraw RectangleCreate(DrawPosition.Left - 3, DrawPosition.Top - 3, DrawPosition.Width + 6, DrawPosition.Height + 6), vbYellow
                End If
                
                DrawButtonWithUI DrawPosition, Spacecrafts(i).Name, "system(""switch_select_spacecraft " & i & """)"
                DrawPosition.Top = DrawPosition.Top + DrawPosition.Height + 10
            End If
        Next
    End With
End Sub

'������Ϸ�ٶȿ��ƽ���
Private Sub DrawSpeedControlBar(Position As Rectangle)
    ForeColor = vbBlack
    FillStyle = 0
    
    With Position
        If GameSpeed <= 0 Then
            DrawButtonWithUI Position, "��ͣ", "", , RGB(128, 128, 128)
        Else
            DrawButtonWithUI Position, "��ͣ", "system(""set_speed 0"")", , RGB(192, 192, 192)
        End If
        .Left = .Left + .Width + 10
        
        If GameSpeed = 1 Then
            DrawButtonWithUI Position, "1��", "", , RGB(128, 128, 128)
        Else
            DrawButtonWithUI Position, "1��", "system(""set_speed 1"")", , RGB(192, 192, 192)
        End If
        .Left = .Left + .Width + 10
        
        If GameSpeed = 2 Then
            DrawButtonWithUI Position, "2��", "", , RGB(128, 128, 128)
        Else
            DrawButtonWithUI Position, "2��", "system(""set_speed 2"")", , RGB(192, 192, 192)
        End If
        .Left = .Left + .Width + 10
        
        If GameSpeed = 4 Then
            DrawButtonWithUI Position, "4��", "", , RGB(128, 128, 128)
        Else
            DrawButtonWithUI Position, "4��", "system(""set_speed 4"")", , RGB(192, 192, 192)
        End If
    End With
End Sub

Private Sub DrawStartGame() '���ƿ�ʼ��Ϸ�˵�
    Dim ButtonPosition As Rectangle

    Font = "΢���ź�"
    ForeColor = vbBlack
    FillStyle = 0
    FontSize = 9
    Font = "����"
    ButtonPosition = RectangleCreate(40, 40, 180, 40)
    DrawButton ButtonPosition, "������Ϸ"
'    AddUIObjectList "Button", ButtonPosition, 1, 3, ShowUILoadGame
    
    DrawButtonWithUI RectangleCreate(40, 100, 180, 40), "�µ���Ϸ", "system(""change_window " & FormGame & """)"
    
    ButtonPosition = RectangleCreate(40, 160, 180, 40)
    DrawButton ButtonPosition, "����"
'    AddUIObjectList ButtonPosition, 1, 3, ShowUINewCampaign
    
    ButtonPosition = RectangleCreate(40, 220, 180, 40)
    DrawButton ButtonPosition, "�̳�"
'    AddUIObjectList ButtonPosition, 1, 3, ShowUITutorial
    
    DrawMenuUI
End Sub

Private Sub DrawToolTip(ByVal X As Long, ByVal Y As Long, ByVal Text As String) '������ʾ��Ϣ
    ForeColor = vbBlack
    FillStyle = 0
    FillColor = vbWhite
    Line (X, Y)-(X + TextWidth(Text) + 10, Y + TextHeight(Text) + 10), , B
    MyPrint Text, X + 5, Y + 5
End Sub

Private Sub DrawToolTips() '��鲢���Ƶ�ǰ������ʾ��Ϣ
    Dim MouseX As Long, MouseY As Long '��ǰ����x��y����
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

'Ч������
Private Sub EffectCalculate(Effect As Effect, Planet As Planet, Efficiency As Double)
    With Effect
        Select Case .Type
        Case EffectOxygen
            PlanetChangeMaterial Planet, "����", .Amont * Efficiency
        Case EffectGas
            PlanetChangeMaterial Planet, "�ȶ�����", .Amont * Efficiency
        Case EffectWater
            PlanetChangeMaterial Planet, "ˮ", .Amont * Efficiency
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

Private Sub ExpansionModule(Planet As Planet, ByVal n As Long) '��������
    Dim i As Long
    
    If Planet.HomeWorld Then
        MsgBox "�����޷�������ĸ��"
        Exit Sub
    End If
    
    With ModuleTypes(Planet.Modules(n).Type)
        '����ʽ��Ƿ��㹻
        If Money < .Cost Then
            MsgBox "ȱ��" & NumberFormat(.Cost - Money) & "�ʽ�"
            Exit Sub
        End If
        
        '���ռ��Ƿ��㹻
        If .Space > 0 Then
            If PlanetGetUsedBlock(Planet) + .Space > Planet.UtilizingBlock Then
                MsgBox Planet.Name & "ȱ��" & NumberFormat(PlanetGetUsedBlock(Planet) + .Space - Planet.UtilizingBlock) & "�����ռ�"
                Exit Sub
            End If
        End If
        
        '�����Դ�Ƿ��㹻
        For i = 1 To UBound(.Resources)
            If Planet.Resources(.Resources(i).Type) < .Resources(i).Amont Then
                MsgBox "ȱ��" & NumberFormat(.Resources(i).Amont - Planet.Resources(.Resources(i).Type)) & GetResourceName(.Resources(i).Type)
                Exit Sub
            End If
        Next
        
        Planet.Modules(n).Construction = 0
    
        '�۳���Դ
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
        GetEffectName = "��Ч��"
    Case EffectOxygen
        GetEffectName = "����"
    Case EffectGas
        GetEffectName = "��ѹ"
    Case EffectWater
        GetEffectName = "ˮ"
    Case EffectTempreture
        GetEffectName = "�¶�"
    Case EffectReflectivity
        GetEffectName = "������"
    Case EffectHousing
        GetEffectName = "ס��"
    Case EffectResource
        GetEffectName = "��Դ"
    Case EffectResearchPoint
        GetEffectName = "�о�����"
    Case EffectPrestige
        GetEffectName = "����"
    Case EffectStorage
        GetEffectName = "�洢�ռ�"
    Case EffectSolarPower
        GetEffectName = "̫����"
    Case EffectRunOff
        GetEffectName = "������"
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

Private Function GetGrowthRate(Planet As Planet) As Double '�������ǵ���Ȼ������
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
        GetResourceName = "����Դ"
    Case ResourceRock
        GetResourceName = "��ʯ"
    Case ResourceMineral
        GetResourceName = "����"
    Case ResourceMetel
        GetResourceName = "����"
    Case ResourceComposites
        GetResourceName = "���ϲ���"
    Case ResourceFood
        GetResourceName = "ʳ��"
    Case ResourceCleanWater
        GetResourceName = "��ˮ"
    Case ResourceRobot
        GetResourceName = "������"
    Case ResourceFurniture
        GetResourceName = "�Ҿ�"
    Case ResourceElectricAppliance
        GetResourceName = "�ҵ�"
    Case ResourceBioMaterial
        GetResourceName = "�������"
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

Private Function GetTooltip(UIObject As UIObject) As String '�����ʾ��Ϣ
    Dim i As Long
    
    GetTooltip = ""
    Select Case UIObject.ClickAction
    Case 0 '��Ч��
    Case 1 '��չ�˿�
        GetTooltip = "����1�˿�"
    Case 2 '���콨��
'        With ModuleTypes(UIObject.ClickAddedCode)
'            GetTooltip = "����" & .Name
'            GetTooltip = GetTooltip & vbCrLf & .Description
'
'            If Money >= .Cost Then
'                GetTooltip = GetTooltip & vbCrLf & "����:" & .Cost
'            Else
'                GetTooltip = GetTooltip & vbCrLf & "����:" & .Cost & "(����)"
'            End If
'            GetTooltip = GetTooltip & vbCrLf & "����ʱ��:" & .BuildTime & "��"
'            If PlanetGetUsedBlock(Planets(SelectPlanet)) + .Space <= Planets(SelectPlanet).UtilizingBlock Then
'                GetTooltip = GetTooltip & vbCrLf & "�����ռ�:" & .Space
'            Else
'                GetTooltip = GetTooltip & vbCrLf & "�����ռ�:" & .Space & "(����)"
'            End If
'
'            For i = 1 To UBound(.Resources)
'                GetTooltip = GetTooltip & vbCrLf & ResourceToString(.Resources(i))
'                With .Resources(i)
'                    If Planets(SelectPlanet).Resources(.Type) < .Amont Then
'                        GetTooltip = GetTooltip & "(����)"
'                    End If
'                End With
'            Next
'        End With
        
    Case 3 'չʾUI
    Case 4 '��ʾ����
'        With Planets(SelectPlanet).Modules(UIObject.ClickAddedCode)
'            GetTooltip = ModuleTypes(.Type).Name & " ��ģ:" & .Size
'            GetTooltip = GetTooltip & vbCrLf & ModuleTypes(.Type).Description
'        End With
    Case 5 '������ҳ
    Case 6 '�������
'        GetTooltip = "���" & ModuleTypes(Planets(SelectPlanet).Modules(UIObject.ClickAddedCode).Type).Name
    Case 7 'һ����������Դ
    Case 8 '���ӳ���������Դ
    Case 9 '���ٳ���������Դ
    Case 10 '�л�����
    Case 11 '���ӽ����ռ�
        GetTooltip = "����1�����ռ�"
    Case 12 'ѡ������
'        If UIObject.ClickAddedCode = 0 Then
'            GetTooltip = "��ʾ��ϵ"
'        Else
'            GetTooltip = Planets(UIObject.ClickAddedCode).Name
'        End If
        
    Case 13 '�ر��¼�
    Case 14 '�ı���Ϸ�ٶ�
    Case 15 '��ʾ/������ͣ����
    Case 16 '��/���ý���
    Case 17 '��������
'        With ModuleTypes(Planets(SelectPlanet).Modules(UIObject.ClickAddedCode).Type)
'            GetTooltip = "����" & .Name
'
'            If Money >= .Cost Then
'                GetTooltip = GetTooltip & vbCrLf & "����:" & .Cost
'            Else
'                GetTooltip = GetTooltip & vbCrLf & "����:" & .Cost & "(����)"
'            End If
'
'            If PlanetGetUsedBlock(Planets(SelectPlanet)) + .Space <= Planets(SelectPlanet).UtilizingBlock Then
'                GetTooltip = GetTooltip & vbCrLf & "�����ռ�:" & .Space
'            Else
'                GetTooltip = GetTooltip & vbCrLf & "�����ռ�:" & .Space & "(����)"
'            End If
'
'            For i = 1 To UBound(.Resources)
'                GetTooltip = GetTooltip & vbCrLf & ResourceToString(.Resources(i))
'                With .Resources(i)
'                    If Planets(SelectPlanet).Resources(.Type) < .Amont Then
'                        GetTooltip = GetTooltip & "(����)"
'                    End If
'                End With
'            Next
'        End With
    Case 18 '�ı���水ť
    Case 19 '����ֿ���Դ
'        GetTooltip = GetResourceName(UIObject.ClickAddedCode) & ":" & NumberFormat(Planets(SelectPlanet).Resources(UIObject.ClickAddedCode)) & "/" & NumberFormat(GetStorage(Planets(SelectPlanet)))
'        GetTooltip = GetTooltip & vbCrLf & "���:" & NumberFormat(Min(-0.01 * (Planets(SelectPlanet).Resources(UIObject.ClickAddedCode) - GetStorage(Planets(SelectPlanet))), 0))
    End Select
End Function

Private Sub LoadEvent()
    ReDim Events(1)
    
    With Events(1)
        .Title = "���"
        .Content = "��ӭ������Ϸ�����ո����ͣ��Ϸ����ESC����ʾ/����" & vbCrLf & "��ͣ�˵�"
        ReDim .Options(1)
        .Options(1) = "ȷ��"
    End With
End Sub

Private Sub LoadMaterial()
    ReDim MaterialTypes(25)
    
    With MaterialTypes(1)
        .Name = "��"
        .HighTemp = 273.15
        .HighTempTarget = "ˮ"
        .State = MaterialStateSolid
    End With
    
    With MaterialTypes(2)
        .Name = "ˮ"
        .LowTemp = 273.15
        .LowTempTarget = "��"
        .HighTemp = 373.15
        .HighTempTarget = "ˮ����"
        .State = MaterialStateLiquid
    End With
    
    With MaterialTypes(3)
        .Name = "ˮ����"
        .LowTemp = 373.15
        .LowTempTarget = "ˮ"
        .State = MaterialStateGas
    End With
    
    With MaterialTypes(4)
        .Name = "��̬��"
        .HighTemp = 54.3
        .HighTempTarget = "Һ̬��"
        .State = MaterialStateSolid
    End With
    
    With MaterialTypes(5)
        .Name = "Һ̬��"
        .LowTemp = 54.3
        .LowTempTarget = "��̬��"
        .HighTemp = 90.15
        .HighTempTarget = "����"
        .State = MaterialStateLiquid
    End With
    
    With MaterialTypes(6)
        .Name = "����"
        .LowTemp = 90.15
        .LowTempTarget = "Һ̬��"
        .State = MaterialStateGas
    End With
    
    With MaterialTypes(7)
        .Name = "��̬��"
        .HighTemp = 14
        .HighTempTarget = "Һ̬��"
        .State = MaterialStateSolid
    End With
    
    With MaterialTypes(8)
        .Name = "Һ̬��"
        .LowTemp = 14
        .LowTempTarget = "��̬��"
        .HighTemp = 21
        .HighTempTarget = "����"
        .State = MaterialStateLiquid
    End With
    
    With MaterialTypes(9)
        .Name = "����"
        .LowTemp = 21
        .LowTempTarget = "Һ̬��"
        .State = MaterialStateGas
    End With
    
    With MaterialTypes(10)
        .Name = "��̬�ȶ�����"
        .HighTemp = 62
        .HighTempTarget = "Һ̬�ȶ�����"
        .State = MaterialStateSolid
    End With
    
    With MaterialTypes(11)
        .Name = "Һ̬�ȶ�����"
        .LowTemp = 62
        .LowTempTarget = "��̬�ȶ�����"
        .HighTemp = 77
        .HighTempTarget = "�ȶ�����"
        .State = MaterialStateLiquid
    End With
    
    With MaterialTypes(12)
        .Name = "�ȶ�����"
        .LowTemp = 77
        .LowTempTarget = "Һ̬�ȶ�����"
        .State = MaterialStateGas
    End With
    
    With MaterialTypes(13)
        .Name = "��̬̼������"
        .HighTemp = 195
        .HighTempTarget = "̼����������"
        .State = MaterialStateSolid
    End With
    
    With MaterialTypes(15)
        .Name = "̼����������"
        .LowTemp = 195
        .LowTempTarget = "��̬̼������"
        .State = MaterialStateGas
    End With
    
    With MaterialTypes(16)
        .Name = "��̬����"
        .HighTemp = 85
        .HighTempTarget = "Һ̬����"
        .State = MaterialStateSolid
    End With
    
    With MaterialTypes(17)
        .Name = "Һ̬����"
        .LowTemp = 85
        .LowTempTarget = "��̬����"
        .HighTemp = 231
        .HighTempTarget = "��̬����"
        .State = MaterialStateLiquid
    End With
    
    With MaterialTypes(18)
        .Name = "��̬����"
        .LowTemp = 231
        .LowTempTarget = "Һ̬����"
        .State = MaterialStateGas
    End With
    
    With MaterialTypes(19)
        .Name = "��ʯ"
        .HighTemp = 1670
        .HighTempTarget = "�ҽ�"
        .State = MaterialStateSolid
    End With
    
    With MaterialTypes(20)
        .Name = "�ҽ�"
        .LowTemp = 1670
        .LowTempTarget = "��ʯ"
        .HighTemp = 2630
        .HighTempTarget = "��̬��"
        .State = MaterialStateLiquid
    End With
    
    With MaterialTypes(21)
        .Name = "��̬��"
        .LowTemp = 2630
        .LowTempTarget = "�ҽ�"
        .State = MaterialStateGas
    End With
    
    With MaterialTypes(22)
        .Name = "����"
        .HighTemp = 1808
        .HighTempTarget = "Һ̬����"
        .State = MaterialStateSolid
    End With
    
    With MaterialTypes(23)
        .Name = "Һ̬����"
        .LowTemp = 1808
        .LowTempTarget = "����"
        .HighTemp = 3023
        .HighTempTarget = "��̬����"
        .State = MaterialStateLiquid
    End With
    
    With MaterialTypes(24)
        .Name = "��̬����"
        .LowTemp = 3023
        .LowTempTarget = "Һ̬����"
        .State = MaterialStateGas
    End With
    
    With MaterialTypes(25)
        .Name = "����"
        .HighTemp = 1670
        .HighTempTarget = "�ҽ�"
        .State = MaterialStateSolid
    End With
End Sub

Private Sub LoadModule()
    Dim Effects() As Effect
    Dim Resources() As Resource
    Dim i As Long
    
    ReDim ModuleTypes(0)
'    'ũ��
'    ModuleTypes(5) = MoudleTypeCreate("ũ��", "����ʳ�", 100, 1, False, -10, 10, 0, 60)
'    With ModuleTypes(5)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceComposites, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectResource
'        .Effects(1).EffectResources = ResourceCreate(ResourceFood, 30)
'    End With
    
    'ũ��(С)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 0.05))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("ũ��(С)", Effects, Resources, "С��ũ�����ʳ�", 100, 0.1, False, -1, 100, -0.02, 60)
    
    'ũ��(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 55))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("ũ��(��)", Effects, Resources, "����ũ�����ʳ�", 100000, 100, False, -1000, 100000, -20, 120)
    
    'ũ��(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 60000))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("ũ��(��)", Effects, Resources, "����ũ�����ʳ�", 100000000, 100000, False, -1000000, 100000000, -20000, 180)

'    '��ͳũҵ
'    ModuleTypes(21) = MoudleTypeCreate("��ͳũҵ", "ʹ�ø��ء�����ʳ��", 100, 8, True, -30, 50000000, -10000000, 300)
'    With ModuleTypes(21)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceComposites, 20000)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectResource
'        .Effects(1).EffectResources = ResourceCreate(ResourceFood, 100000000)
'    End With
    
    '��ͳũҵ(С)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(0)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 0.05))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("��ͳũҵ(С)", Effects, Resources, "С�ʹ�ͳũҵ��ʹ�ô�������������ʳ�", 0.1, 1, True, 0, 100, 0, 60)
    
    '��ͳũҵ(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(0)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 55))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("��ͳũҵ(��)", Effects, Resources, "���ʹ�ͳũҵ��ʹ�ô�������������ʳ�", 100, 1000, True, 0, 100000, 0, 120)
    
    '��ͳũҵ(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(0)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 60000))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("��ͳũҵ(��)", Effects, Resources, "���ʹ�ͳũҵ��ʹ�ô�������������ʳ�", 100000, 1000000, True, 0, 100000000, 0, 180)
    
    '�ִ�ũҵ(С)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 0.05))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("�ִ�ũҵ(С)", Effects, Resources, "С���ִ�ũҵ��ʹ�ø��أ�����ʳ�", 10, 10, True, -0.8, 100, -0.1, 100)
    
    '�ִ�ũҵ(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 55))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("�ִ�ũҵ(��)", Effects, Resources, "�����ִ�ũҵ��ʹ�ø��أ�����ʳ�", 10000, 10000, True, -800, 100000, -100, 200)
    
    '�ִ�ũҵ(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceFood, 60000))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("�ִ�ũҵ(��)", Effects, Resources, "�����ִ�ũҵ��ʹ�ø��أ�����ʳ�", 10000000, 10000000, True, -800000, 100000000, -100000, 300)
    
'    '�ƿ�վ
'    ModuleTypes(1) = MoudleTypeCreate("�ƿ�վ", "�Ը������С��վ�㡣���Բ����������е�����", 500, 10, False, -50, 0, 0, 60)
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

    '�ƿ�վ(С)
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
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("�ƿ�վ(С)", Effects, Resources, "�Ը������С��վ�㡣���Բ����������е�����", 10, 10, False, -0.8, 100, -0.1, 100)
    
    '�ƿ�վ(��)
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
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("�ƿ�վ(��)", Effects, Resources, "�Ը����������վ�㡣���Բ����������е�����", 10000, 10000, False, -800, 100000, -100, 200)
    
    '�ƿ�վ(��)
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
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("�ƿ�վ(��)", Effects, Resources, "�Ը�����Ĵ���վ�㡣���Բ����������е�����", 10000000, 10000000, False, -800000, 100000000, -100000, 300)
    
'    '̫���ܵ�ذ�
'    ModuleTypes(3) = MoudleTypeCreate("̫���ܵ�ذ�", "�ṩ��������Ҫ̫���⡣", 100, 1, False, -10, 0, 100, 60)
'    With ModuleTypes(3)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceComposites, 10)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectSolarPower
'    End With

    '̫���ܷ���վ(С)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectSolarPower, ResourceNull)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("̫���ܷ���վ(С)", Effects, Resources, "С��̫���ܷ���վ���ṩ��������Ҫ���ա�", 100, 10, False, -1, 100, 0.1, 60)
    
    '̫���ܷ���վ(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectSolarPower, ResourceNull)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("̫���ܷ���վ(��)", Effects, Resources, "����̫���ܷ���վ���ṩ��������Ҫ���ա�", 100000, 10000, False, -1000, 100000, 100, 120)
    
    '̫���ܷ���վ(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectSolarPower, ResourceNull)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("̫���ܷ���վ(��)", Effects, Resources, "����̫���ܷ���վ���ṩ��������Ҫ���ա�", 100000000, 10000000, False, -1000000, 100000000, 100000, 180)

'    '��������վ
'    ModuleTypes(22) = MoudleTypeCreate("��������վ", "���ÿ����������", 1000, 1, False, -300, 100, 1000, 60)
'    With ModuleTypes(22)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceMetel, 100)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectResource
'        .Effects(1).EffectResources = ResourceCreate(ResourceMineral, -100)
'    End With

    '��������վ(С)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceMineral, -0.05))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("��������վ(С)", Effects, Resources, "С�ͻ�������վ�����ÿ������������", 100, 50, False, -1, 100, 0.1, 60)
    
    '��������վ(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectSolarPower, ResourceCreate(ResourceMineral, -50))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("��������վ(��)", Effects, Resources, "���ͻ�������վ�����ÿ������������", 100000, 50000, False, -1000, 100000, 100, 120)
    
    '��������վ(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectSolarPower, ResourceCreate(ResourceMineral, -50000))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("��������վ(��)", Effects, Resources, "���ͻ�������վ�����ÿ������������", 100000000, 50000000, False, -1000000, 100000000, 100000, 180)

'    '��סģ��
'    ModuleTypes(2) = MoudleTypeCreate("��סģ��", "�ṩסլ�ռ䣬��Ҫ����ά�֡�", 100, 1, False, -10, 0, -10, 60)
'    With ModuleTypes(2)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceComposites, 10)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectHousing
'        .Effects(1).Amont = 100
'    End With

    '��סģ��(С)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(2)
    Resources(1) = ResourceCreate(ResourceComposites, 10)
    Resources(2) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectHousing, ResourceNull, 1000)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("��סģ��(С)", Effects, Resources, "С�;�סģ�顣�ṩסլ�ռ䣬��Ҫ����ά�֡�", 100, 0.1, False, -1, 0, -0.02, 60)
    
    '��סģ��(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(2)
    Resources(1) = ResourceCreate(ResourceComposites, 10000)
    Resources(2) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectHousing, ResourceNull, 1000000)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("��סģ��(��)", Effects, Resources, "���;�סģ�顣�ṩסլ�ռ䣬��Ҫ����ά�֡�", 100000, 100, False, -1000, 0, -20, 120)
    
    '��סģ��(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(2)
    Resources(1) = ResourceCreate(ResourceComposites, 10000000)
    Resources(2) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectHousing, ResourceNull, 1000000)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("��סģ��(��)", Effects, Resources, "���;�סģ�顣�ṩסլ�ռ䣬��Ҫ����ά�֡�", 100000000, 100000, False, -1000000, 0, -20000, 180)

'    'ó�׹�˾
'    ModuleTypes(4) = MoudleTypeCreate("ó�׹�˾", "�����ʽ�", 100, 1, False, 50, 10, -10, 60)
'    With ModuleTypes(4)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceComposites, 20)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectNone
'    End With

    'ó�׹�˾(С)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 10)
    ReDim Effects(1)
    Effects(1) = EffectNull
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("ó�׹�˾(С)", Effects, Resources, "С��ó�׹�˾�������ʽ�", 100, 0.1, False, 5, 100, -0.01, 60)
    
    'ó�׹�˾(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 10000)
    ReDim Effects(1)
    Effects(1) = EffectNull
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("ó�׹�˾(��)", Effects, Resources, "����ó�׹�˾�������ʽ�", 100000, 100, False, 5000, 100000, -10, 120)
    
    'ó�׹�˾(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectNull
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("ó�׹�˾(��)", Effects, Resources, "����ó�׹�˾�������ʽ�", 100000000, 100000, False, 5000000, 100000000, -10000, 180)
    
'    'ˮ������
'    ModuleTypes(6) = MoudleTypeCreate("ˮ������", "��������ˮ��", 100, 1, False, -30, 0, -30, 60)
'    With ModuleTypes(6)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceComposites, 20)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectResource
'        .Effects(1).EffectResources = ResourceCreate(ResourceCleanWater, 30)
'    End With

    'ˮ������(С)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 20)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceCleanWater, 0.1))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("ˮ������(С)", Effects, Resources, "С��ˮ����������������ˮ��", 50, 0.1, False, -1, 100, -0.02, 60)
    
    'ˮ������(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 20000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceCleanWater, 100))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("ˮ������(��)", Effects, Resources, "����ˮ����������������ˮ��", 50000, 100, False, -1000, 100000, -20, 120)
    
    'ˮ������(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceComposites, 20000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceCleanWater, 100000))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("ˮ������(��)", Effects, Resources, "����ˮ����������������ˮ��", 50000000, 100000, False, -1000000, 100000000, -20000, 180)

'    '��
'    ModuleTypes(7) = MoudleTypeCreate("��", "����������", 100, 1, False, -10, 10, -30, 60)
'    With ModuleTypes(7)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceMetel, 20)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectResource
'        .Effects(1).EffectResources = ResourceCreate(ResourceMetel, 10)
'    End With

    '��(С)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceMetel, 0.1))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("��(С)", Effects, Resources, "С�Ϳ󳡡�����������", 20, 10, False, -1, 100, -0.02, 60)
    
    '��(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceMetel, 100))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("��(��)", Effects, Resources, "���Ϳ󳡡�����������", 20000, 10000, False, -1000, 100000, -20, 120)
    
    '��(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectResource, ResourceCreate(ResourceMetel, 100000))
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("��(��)", Effects, Resources, "���Ϳ󳡡�����������", 20000000, 10000000, False, -1000000, 100000000, -20000, 180)

'    '���巢����
'    ModuleTypes(8) = MoudleTypeCreate("���巢����", "��������ͷ����塣", 100, 1, False, -10, 10, -10, 60)
'    With ModuleTypes(8)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceRock, 10)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectGas
'        .Effects(1).Amont = 0.00001
'    End With

    '���巢����(С)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectGas, ResourceNull, 0.00000001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("���巢����(С)", Effects, Resources, "С�����巢��������������ͷ����塣", 20, 0.1, False, -1, 100, -0.02, 60)
    
    '���巢����(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectGas, ResourceNull, 0.00001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("���巢����(��)", Effects, Resources, "�������巢��������������ͷ����塣", 20000, 100, False, -1000, 100000, -20, 120)
    
    '���巢����(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectGas, ResourceNull, 0.01)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("���巢����(��)", Effects, Resources, "�������巢��������������ͷ����塣", 20000000, 100000, False, -1000000, 100000000, -20000, 180)

'    '���崢��
'    ModuleTypes(9) = MoudleTypeCreate("���崢��", "���մ����е����塣", 100, 1, False, -10, 10, -10, 60)
'    With ModuleTypes(9)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceRock, 10)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectGas
'        .Effects(1).Amont = -0.00001
'    End With

    '���崢��(С)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectGas, ResourceNull, -0.00000001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("���崢��(С)", Effects, Resources, "С�����崢�ޡ����մ����е����塣", 20, 0.1, False, -1, 100, -0.02, 60)
    
    '���崢��(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectGas, ResourceNull, -0.00001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("���崢��(��)", Effects, Resources, "�������崢�ޡ����մ����е����塣", 20000, 100, False, -1000, 100000, -20, 120)
    
    '���崢��(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectGas, ResourceNull, -0.01)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("���崢��(��)", Effects, Resources, "�������崢�ޡ����մ����е����塣", 20000000, 100000, False, -1000000, 100000000, -20000, 180)

'    '����ˮ��
'    ModuleTypes(10) = MoudleTypeCreate("����ˮ��", "���ߺ�ƽ�档", 100, 1, False, -10, 10, -10, 60)
'    With ModuleTypes(10)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceRock, 10)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectWater
'        .Effects(1).Amont = 0.005
'    End With

    '����ˮ��(С)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectWater, ResourceNull, 0.000001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("����ˮ��(С)", Effects, Resources, "С�͵���ˮ�󾮡����ߺ�ƽ�档", 20, 0.1, False, -1, 100, -0.02, 60)
    
    '����ˮ��(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectWater, ResourceNull, 0.001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("����ˮ��(��)", Effects, Resources, "���͵���ˮ�󾮡����ߺ�ƽ�档", 20000, 100, False, -1000, 100000, -20, 120)
    
    '����ˮ��(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectWater, ResourceNull, 1)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("����ˮ��(��)", Effects, Resources, "���͵���ˮ�󾮡����ߺ�ƽ�档", 20000000, 100000, False, -1000000, 100000000, -20000, 180)

'    'ˮ�̻�ϵͳ
'    ModuleTypes(11) = MoudleTypeCreate("ˮ�̻�ϵͳ", "���ͺ�ƽ�档", 100, 1, False, -10, 10, -10, 60)
'    With ModuleTypes(11)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceRock, 10)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectWater
'        .Effects(1).Amont = -0.005
'    End With

    'ˮ�̻�ϵͳ(С)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectWater, ResourceNull, -0.000001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("ˮ�̻�ϵͳ(С)", Effects, Resources, "С��ˮ�̻�ϵͳ�����ͺ�ƽ�档", 20, 0.1, False, -1, 100, -0.02, 60)
    
    'ˮ�̻�ϵͳ(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectWater, ResourceNull, -0.001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("ˮ�̻�ϵͳ(��)", Effects, Resources, "����ˮ�̻�ϵͳ�����ͺ�ƽ�档", 20000, 100, False, -1000, 100000, -20, 120)
    
    'ˮ�̻�ϵͳ(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectWater, ResourceNull, -1)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("ˮ�̻�ϵͳ(��)", Effects, Resources, "����ˮ�̻�ϵͳ�����ͺ�ƽ�档", 20000000, 100000, False, -1000000, 100000000, -20000, 180)

'    '��ʯ�Ƚ�ϵͳ
'    ModuleTypes(12) = MoudleTypeCreate("��ʯ�Ƚ�ϵͳ", "��������ͷ�������", 100, 1, False, -10, 10, -10, 60)
'    With ModuleTypes(12)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceRock, 20)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectOxygen
'        .Effects(1).Amont = 0.00001
'    End With

    '��ʯ�Ƚ�ϵͳ(С)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectOxygen, ResourceNull, 0.00000001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("��ʯ�Ƚ�ϵͳ(С)", Effects, Resources, "С����ʯ�Ƚ�ϵͳ����������ͷ�������", 20, 0.1, False, -1, 100, -0.02, 60)
    
    '��ʯ�Ƚ�ϵͳ(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectOxygen, ResourceNull, 0.00001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("��ʯ�Ƚ�ϵͳ(��)", Effects, Resources, "������ʯ�Ƚ�ϵͳ����������ͷ�������", 20000, 100, False, -1000, 100000, -20, 120)
    
    '��ʯ�Ƚ�ϵͳ(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectOxygen, ResourceNull, 0.01)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("��ʯ�Ƚ�ϵͳ(��)", Effects, Resources, "������ʯ�Ƚ�ϵͳ����������ͷ�������", 20000000, 100000, False, -1000000, 100000000, -20000, 180)

'    '�����̻���
'    ModuleTypes(13) = MoudleTypeCreate("�����̻���", "��ȥ�������е�������", 100, 1, False, -10, 10, -10, 60)
'    With ModuleTypes(13)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceRock, 10)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectOxygen
'        .Effects(1).Amont = -0.00001
'    End With

    '����������(С)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectOxygen, ResourceNull, 0.00000001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("����������(С)", Effects, Resources, "С����������������ȥ�������е�������", 20, 0.1, False, -1, 100, -0.02, 60)
    
    '����������(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectOxygen, ResourceNull, 0.00001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("����������(��)", Effects, Resources, "������������������ȥ�������е�������", 20000, 100, False, -1000, 100000, -20, 120)
    
    '����������(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectOxygen, ResourceNull, 0.01)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("����������(��)", Effects, Resources, "������������������ȥ�������е�������", 20000000, 100000, False, -1000000, 100000000, -20000, 180)

'    '���Ǽ�����
'    ModuleTypes(14) = MoudleTypeCreate("���Ǽ�����", "���������¶ȡ�", 100, 1, False, -10, 10, -10, 60)
'    With ModuleTypes(14)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceRock, 10)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectTempreture
'        .Effects(1).Amont = 0.5
'    End With

    '���Ǽ�����(С)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectTempreture, ResourceNull, 0.0001)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("���Ǽ�����(С)", Effects, Resources, "С�����Ǽ����������ͺ�ƽ�档", 20, 0.1, False, -1, 100, -0.02, 60)
    
    '���Ǽ�����(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectWater, ResourceNull, 0.1)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("���Ǽ�����(��)", Effects, Resources, "�������Ǽ����������ͺ�ƽ�档", 20000, 100, False, -1000, 100000, -20, 120)
    
    '���Ǽ�����(��)
    ReDim Preserve ModuleTypes(UBound(ModuleTypes) + 1)
    ReDim Resources(1)
    Resources(1) = ResourceCreate(ResourceMetel, 10000000)
    ReDim Effects(1)
    Effects(1) = EffectCreate(EffectWater, ResourceNull, 100)
    ModuleTypes(UBound(ModuleTypes)) = MoudleTypeCreate("���Ǽ�����(��)", Effects, Resources, "�������Ǽ����������ͺ�ƽ�档", 20000000, 100000, False, -1000000, 100000000, -20000, 180)

'    '�����˹���
'    ModuleTypes(15) = MoudleTypeCreate("�����˹���", "ʹ�ý�����������ˡ�", 100, 1, False, -10, 10, -30, 60)
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
'    '�о���
'    ModuleTypes(16) = MoudleTypeCreate("�о���", "�����о���", 100, 1, False, -10, 10, -10, 60)
'    With ModuleTypes(16)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceComposites, 20)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectResearchPoint
'        .Effects(1).Amont = 1
'    End With
'
'    '����̨
'    ModuleTypes(17) = MoudleTypeCreate("����̨", "����������", 100, 1, False, -10, 10, -10, 60)
'    With ModuleTypes(17)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceComposites, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectPrestige
'        .Effects(1).Amont = 1
'    End With
'
'    '�ֿ�
'    ModuleTypes(18) = MoudleTypeCreate("�ֿ�", "������Ʒ������", 100, 1, False, 0, 10, 0, 60)
'    With ModuleTypes(18)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectStorage
'        .Effects(1).Amont = 500
'    End With
'
'    '����ˮ��
'    ModuleTypes(19) = MoudleTypeCreate("����ˮ��", "����Ȼˮ���л�ȡˮ��������ˮ��", 10000000, 10, False, -1000000, 10000000, -10000000, 300)
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
'    '������
'    ModuleTypes(20) = MoudleTypeCreate("������", "��������ҵ(δ���)��", 100, 1, False, -30, 10, -30, 60)
'    With ModuleTypes(20)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceComposites, 20)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectResource
'        .Effects(1).EffectResources = ResourceCreate(ResourceCleanWater, 10)
'    End With
'
'    '��ͳס��
'    ModuleTypes(23) = MoudleTypeCreate("��ͳס��", "�ṩס��", 10000000, 10, True, -10000000, 10000000, -10000000, 300)
'    With ModuleTypes(23)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceMetel, 20000)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectHousing
'        .Effects(1).Amont = 1000000000
'    End With
'
'    '���������
'    ModuleTypes(24) = MoudleTypeCreate("���������", "��������ķ�����", 100, 1, False, -10, 0, 0, 60)
'    With ModuleTypes(24)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceComposites, 10)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectReflectivity
'        .Effects(1).Amont = 0.99
'    End With
'
'    '������������
'    ModuleTypes(25) = MoudleTypeCreate("������������", "ʹ��ȼ�Ϸ���", 1000, 1, True, -3000000, 1000000, 10000000, 300)
'    With ModuleTypes(25)
'        ReDim .Resources(2)
'        .Resources(1) = ResourceCreate(ResourceComposites, 10000)
'        .Resources(2) = ResourceCreate(ResourceMetel, 10000)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectResource
'        .Effects(1).EffectResources = ResourceCreate(ResourceMineral, -1000000)
'    End With
'
'    '��ͳ��
'    ModuleTypes(26) = MoudleTypeCreate("��ͳ��", "�������", 10000000, 1, False, -1000000, 50000000, -10000000, 300)
'    With ModuleTypes(26)
'        ReDim .Resources(1)
'        .Resources(1) = ResourceCreate(ResourceMetel, 20000)
'        ReDim .Effects(1)
'        .Effects(1).Type = EffectResource
'        .Effects(1).EffectResources = ResourceCreate(ResourceMineral, 10000000)
'    End With
End Sub

                    
'���غ�����
Private Sub LoadSpaceCraft()
    Dim SpacePosition As SpacePosition
    Dim Effects() As Effect
    Dim Resources() As Resource
    
    ReDim Spacecrafts(1)
    With Spacecrafts(1)
        .Name = "ϣ����(ֳ��)"
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
        .Name = "̫��"
        .Tag = "SUN"
        .Color = vbYellow
        .Magnitude = 4.83
        .Mass = 198900000000# '���ڶ�
        .Type = StarTypeMainSequence
    End With
    
    ReDim Planets(4)
    'ˮ��
    With Planets(1)
        .Name = "ˮ��"
        .Tag = "SUN1"
        .Color = RGB(192, 192, 192)
        .Radio = 2440 'ǧ��
'            .Water = 0 '���ڶ�
'            .Mass = 33011  '���ڶ�
        .Tempreture = 452 '������
'            .Oxygen = 0 '���ڶ�
        .Reflectivity = 0.119
        .RotationPeriod = 58.646 '��
        .OrbitRadius = 0.5791 '��ǧ��
        ReDim .Materials(2)
        .Materials(1).Type = "����"
        .Materials(1).Mass = 17811
        .Materials(2).Type = "��ʯ"
        .Materials(2).Mass = 15200
    End With
    '����
    With Planets(2)
        .Name = "����"
        .Tag = "SUN2"
        .Color = RGB(255, 192, 72)
        .Radio = 6052 'ǧ��
'            .Water = 0 '���ڶ�
'            .Mass = 486750  '���ڶ�
        .Tempreture = 737 '������
'            .Oxygen = 0 '���ڶ�
        .Reflectivity = 0.75
'            .Gas = 48.27 '���ڶ�
        .RotationPeriod = 243 '��
        .OrbitRadius = 1.082 '��ǧ��
        ReDim .Materials(4)
        .Materials(1).Type = "����"
        .Materials(1).Mass = 97750
        .Materials(2).Type = "��ʯ"
        .Materials(2).Mass = 389000
        .Materials(3).Type = "̼����������"
        .Materials(3).Mass = 46.82
        .Materials(4).Type = "�ȶ�����"
        .Materials(4).Mass = 1.45
    End With
    '����
    With Planets(3)
        .Name = "����"
        .Tag = "SUN3"
        .Color = vbBlue
        .Radio = 6371 'ǧ��
'            .Water = 136 '���ڶ�
'            .Mass = 597237  '���ڶ�
        .Tempreture = 289 '������
'            .Oxygen = 0.11885 '���ڶ�
        .Reflectivity = 0.29
'            .Gas = 0.5136 - 0.11885 '���ڶ�
        .RotationPeriod = 1 '��
        .OrbitRadius = 1.496 '��ǧ��
        ReDim .Materials(5)
        .Materials(1).Type = "����"
        .Materials(1).Mass = 147237
        .Materials(2).Type = "��ʯ"
        .Materials(2).Mass = 450000
        .Materials(3).Type = "ˮ"
        .Materials(3).Mass = 136
        .Materials(4).Type = "����"
        .Materials(4).Mass = 0.11885
        .Materials(5).Type = "�ȶ�����"
        .Materials(5).Mass = 0.39475
        
        .HomeWorld = True
        .Colonys.Population = 7000000000#
        
    End With
    '����
    With Planets(4)
        .Name = "����"
        .Tag = "SUN4"
        .Color = vbRed
        .Radio = 3389 'ǧ��
'            .Water = 0.0000001 '���ڶ�
'            .Mass = 64171  '���ڶ�
        .Tempreture = 218 '������
'            .Oxygen = 0 '���ڶ�
        .Reflectivity = 0.16
'            .Gas = 0.0025 '���ڶ�
        .RotationPeriod = 1.0259 '��
        .OrbitRadius = 2.279 '��ǧ��
        ReDim .Materials(5)
        .Materials(1).Type = "����"
        .Materials(1).Mass = 6471
        .Materials(2).Type = "��ʯ"
        .Materials(2).Mass = 57700
        .Materials(3).Type = "��"
        .Materials(3).Mass = 0.0000001
        .Materials(4).Type = "̼����������"
        .Materials(4).Mass = 0.002375
        .Materials(5).Type = "�ȶ�����"
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
        .Name = "̫��ϵ"
        ReDim .Stars(1)
        .Stars(1) = "SUN"
        
        ReDim .Planets(4)
        .Planets(1) = "SUN1" 'ˮ��
        .Planets(2) = "SUN2" '����
        .Planets(3) = "SUN3" '����
        .Planets(4) = "SUN4" '����
    End With
End Sub

Private Sub LoadTechnology()
    ReDim Technologys(5)
    
    '�˾۱�
    With Technologys(1)
        .Name = "�˾۱�"
        .NeedPoints = 1000
        .IsResearched = False
    End With
    
    '�˹�����
    With Technologys(2)
        .Name = "�˹�����"
        .NeedPoints = 100
        .IsResearched = False
    End With
    
    '�ǹ��ٷɴ�
    With Technologys(3)
        .Name = "�ǹ��ٷɴ�"
        .NeedPoints = 200
        .IsResearched = False
    End With
    
    '���ͽ���
    With Technologys(4)
        .Name = "���ͽ���"
        .NeedPoints = 2000
        .IsResearched = False
    End With
    
    'ҽѧ
    With Technologys(5)
        .Name = "ҽѧ"
        .NeedPoints = 100
        .IsResearched = False
    End With
    
    '�߻�
    With Technologys(6)
        .Name = "�߻�"
        .NeedPoints = 100
        .IsResearched = False
    End With
    
    '�������
    With Technologys(6)
        .Name = "�������"
        .NeedPoints = 100
        .IsResearched = False
    End With
    
    '������ʵ
    With Technologys(6)
        .Name = "������ʵ"
        .NeedPoints = 100
        .IsResearched = False
    End With
End Sub

''�����г��۸�
'Private Sub MarketCalculatePrice(Market As Market)
'
'    Market
'    MarketCalculatePrice = Market.Money * Market.Prices
'End Sub

'��ȡ�г��˻�
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
        MyPrint "��ģ:" & .Size, X + 10, CurrentY
        MyPrint "��Դ:" & ModuleTypes(.Type).Power * .Size, X + 10, CurrentY
        MyPrint "�ʽ�:" & ModuleTypes(.Type).Maintenance * .Size, X + 10, CurrentY
        MyPrint "��λ:" & ModuleTypes(.Type).Staff * .Size, X + 10, CurrentY
        
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
        
        MyPrint "ռ�ÿռ�:" & ModuleTypes(.Type).Space * .Size, X + 10, CurrentY
        MyPrint "Ч��:" & Format(.Efficiency, "0.00%"), X + 10, CurrentY
        
        If .Construction < 1 Then
            MyPrint "������", X + 10, CurrentY
            DrawProgressBar RectangleCreate(X + 10, CurrentY + 2, 100, 14), .Construction, RGB(120, 255, 120)
        End If
        
        MyPrint "��߳����¶�" & Format(ModuleTypes(.Type).MaxTempreture - 273.15, "0.00") & "��", X + 10, CurrentY + 2
        MyPrint "������ѹ��" & Format(ModuleTypes(.Type).MaxPressure / 1000000, "0.00") & "MPa", X + 10, CurrentY
        
        MyPrint "�ִ�:", X + 10, CurrentY
        
        If UBound(.Storage) = 0 Then
            MyPrint "��", X + 10, CurrentY
        Else
            For i = 1 To UBound(.Storage)
                MyPrint ResourceToString(.Storage(i)), X + 10, CurrentY
            Next
        End If
        
        If Not Planets(SelectPlanet).HomeWorld Then
            DrawButtonWithUI RectangleCreate(X + 10, Y + 220, 60, 30), "���", "system(""dismantle_module " & Planets(SelectPlanet).Tag & " " & SelectModule & """)"
            
            If Module.Construction = 1 Then
                '������/���ð�ť
                If Module.Enabled Then
                    DrawText = "����"
                Else
                    DrawText = "����"
                End If
                DrawButtonWithUI RectangleCreate(X + 80, Y + 220, 60, 30), DrawText, "system(""switch_module_enabled " & SelectPlanet & " " & SelectModule & """)"
                
                DrawButtonWithUI RectangleCreate(X + 10, Y + 260, 60, 30), "����", "system(""expand_module " & SelectPlanet & " " & SelectModule & """)" '����������ť
            End If
        End If
    End With
    
    DrawCloseButton X + 170, Y + 5, "system(""clear_select"")"
End Sub

'ģ��Ч������
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

'ģ��Ч�ʼ���
Private Sub MoudleEfficiencyCalculate(Module As Module, Planet As Planet, ByVal ID As Long)
    Dim GetModuleType As ModuleType
    Dim GetResources As Resource
    Dim i As Long
    
    With Module
        '����Ч������
        .EfficiencyModifier = 1
        
        If .Enabled = False Then
            .EfficiencyModifier = 0
        End If
        
        '̫����Ч������
        For i = 1 To UBound(ModuleTypes(.Type).Effects)
            If ModuleTypes(.Type).Effects(i).Type = EffectSolarPower Then
                .EfficiencyModifier = .EfficiencyModifier * PlanetGetSolarPower(Stars(1), Planet)
            End If
        Next
        
        If .EfficiencyModifier < 0 Then
            .EfficiencyModifier = 0
        End If
        .Efficiency = .EfficiencyModifier
        
        'Ч������������ϣ�����Ч��
        '����
        If ModuleTypes(.Type).Power < 0 Then
            If .Efficiency > .EfficiencyModifier * GetPowerAdyquacy(Planet) Then
                .Efficiency = .EfficiencyModifier * GetPowerAdyquacy(Planet)
            End If
        End If
        
        '�˿�
        If ModuleTypes(.Type).Staff > 0 Then
            If .Efficiency > .EfficiencyModifier * GetStaffAdyquacy(Planet) Then
                .Efficiency = .EfficiencyModifier * GetStaffAdyquacy(Planet)
            End If
        End If
        
        '��ԴЧ������
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

Private Function MousePosition() As PointApi '��ȡ���λ��
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
    
    'Mode0Ϊ�����ӡ���Ͻ�λ��
    'Mode1Ϊ�����ӡ����λ��
    If Mode = 1 Then
        X = X - 0.5 * TextWidth(Text)
        Y = Y - 0.5 * TextHeight(Text)
    End If
    
    '���зָÿһ�ж������������ٴ�ӡ
    CurrentY = Y
    Lines = Split(Text, vbCrLf)
    For i = LBound(Lines) To UBound(Lines)
        CurrentX = X
        Print Lines(i)
    Next
End Sub

'���ָ�ʽ������
Private Function NumberFormat(n As Variant) As String
    If Abs(n) <= 10000 Then
        NumberFormat = Format(n, "0")
    Else
        If Abs(n) <= 100000000 Then
            NumberFormat = Format(n / 10000, "0.0") & "��"
        Else
            NumberFormat = Format(n / 100000000, "0.0") & "��"
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

'���������Դ
Private Sub PlanetAddResource(Planet As Planet, ResourceType As ResourceEnum, Amont As Double)
    Planet.Resources(ResourceType) = Planet.Resources(ResourceType) + Amont
    If Planet.Resources(ResourceType) < 0 Then Planet.Resources(ResourceType) = 0
End Sub

Private Sub PlanetBuildModule(Planet As Planet, ByVal MoudleType As Long) '���콨��
    Dim i As Long
    
    '����Ƿ�����ĸ��
    If Planet.HomeWorld Then
        MsgBox "�����޷�������ĸ��"
        Exit Sub
    End If
    
    With ModuleTypes(MoudleType)
        '��齨���˾������Ƿ�����
        If Not PlanetIsLiveable(Planet) And .LivableRequire Then
            MsgBox "����ý�����Ҫ�����˾�"
            Exit Sub
        End If
    
        '����ʽ��Ƿ��㹻
        If Money < .Cost Then
            MsgBox "ȱ��" & NumberFormat(.Cost - Money) & "�ʽ�"
            Exit Sub
        End If
        
        '���ռ��Ƿ��㹻
        If .Space > 0 Then
            If PlanetGetUsedBlock(Planet) + .Space > Planet.UtilizingBlock Then
                MsgBox Planet.Name & "ȱ��" & NumberFormat(PlanetGetUsedBlock(Planet) + .Space - Planet.UtilizingBlock) & "�����ռ�"
                Exit Sub
            End If
        End If
        
        '�����Դ�Ƿ��㹻
        For i = 1 To UBound(.Resources)
            If Planet.Resources(.Resources(i).Type) < .Resources(i).Amont Then
                MsgBox "ȱ��" & NumberFormat(.Resources(i).Amont - Planet.Resources(.Resources(i).Type)) & GetResourceName(.Resources(i).Type)
                Exit Sub
            End If
        Next
    
        PlanetAddMoudle Planet, MoudleType
    
        '�۳���Դ
        Money = Money - .Cost
        For i = 1 To UBound(.Resources)
            PlanetAddResource Planet, .Resources(i).Type, -.Resources(i).Amont
        Next
    End With
End Sub

'�ı������ϵ�ĳ�����ʺ���
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

'ɾ������������ĳ������
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

'ɾ��ģ��
Private Sub PlanetDeleteMoudle(Planet As Planet, ID As Long)
    Dim i As Long
    For i = ID To UBound(Planet.Modules) - 1
        Planet.Modules(i) = Planet.Modules(i + 1)
    Next
    ReDim Preserve Planet.Modules(UBound(Planet.Modules) - 1)
End Sub

'��������
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

'��ȡ����������
Private Function PlanetGetEvaporation(Planet As Planet) As Double
    PlanetGetEvaporation = 200000000000# * (Planet.Tempreture / 289) * (Min(PlanetGetWater(Planet), 1) / 0.718) '* (PlanetGetSolarPower(Star, Planet) / 0.72)
End Function

'��ȡ��������ЧӦ
Private Function PlanetGetGreenhouseEffect(Planet As Planet) As Double
    PlanetGetGreenhouseEffect = Tanh(0.41 * Log(PlanetGetPressure(Planet) / 100000 + 1))
End Function

'��ȡ���ǵ�IDֵ
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
            If Planet.Materials(i).Type = "����" Then OxygenMass = OxygenMass + Planet.Materials(i).Mass
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
            If Planet.Materials(i).Type = "ˮ" Then WaterMass = WaterMass + Planet.Materials(i).Mass
        Next
    End With
    PlanetGetWater = WaterMass / 136 * 0.718
End Function

'�����Ƿ��˾�
Private Function PlanetIsLiveable(Planet As Planet) As Boolean
    PlanetIsLiveable = True
    
    '���������ݣ��ж��Ƿ��˾�
    If PlanetGetOxygen(Planet) < 0.18 Or PlanetGetOxygen(Planet) > 0.24 Then PlanetIsLiveable = False
    If PlanetGetWater(Planet) < 0.25 Or PlanetGetWater(Planet) > 0.75 Then PlanetIsLiveable = False
    If Planet.Tempreture < 237 Or Planet.Tempreture > 337 Then PlanetIsLiveable = False
    If PlanetGetPressure(Planet) < 50000 Or PlanetGetPressure(Planet) > 150000 Then PlanetIsLiveable = False
End Function

''���񸶳����ң���������
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

'����ƽ��
Private Function RectangleTranslation(Rect As Rectangle, Optional ByVal X As Long = 0, Optional ByVal Y As Long = 0) As Rectangle
    With RectangleTranslation
        .Left = Rect.Left + X
        .Top = Rect.Top + Y
        .Width = Rect.Width + X
        .Height = Rect.Height + Y
    End With
End Function

'���ƾ���
Private Sub RectangleDraw(Rectangle As Rectangle, Optional ByVal RectangleFillColor = vbWhite)
    FillColor = RectangleFillColor
    Line (Rectangle.Left, Rectangle.Top)-(Rectangle.Left + Rectangle.Width - 1, Rectangle.Top + Rectangle.Height - 1), , B
End Sub

'�жϸ������Ƿ��ھ��ε���ʵ������(���߽�)
Private Function RectangleIsIn(Rect As Rectangle, ByVal X As Long, ByVal Y As Long) As Boolean
    With Rect
        RectangleIsIn = X >= .Left And X <= .Left + .Width And Y >= .Top And Y <= .Top + .Height
    End With
End Function

'��������
Private Function RectangleCreate(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long) As Rectangle
    With RectangleCreate
        .Left = Left
        .Top = Top
        .Width = Width
        .Height = Height
    End With
End Function

'������Դ
Private Function ResourceCreate(ByVal ResourceType As ResourceEnum, ByVal Amont As Double) As Resource
    With ResourceCreate
        .Type = ResourceType
        .Amont = Amont
    End With
End Function

'��������Դ
Private Function ResourceNull() As Resource
    With ResourceNull
        .Type = ResourceNone
        .Amont = 0
    End With
End Function

'�ԡ���Դ���ƣ���Դ��������ʽ�����ַ���
Private Function ResourceToString(Resource As Resource) As String
    ResourceToString = GetResourceName(Resource.Type) & ":" & NumberFormat(Resource.Amont)
End Function

Private Function SaveGame() As Long '�洢��Ϸ
    If Dir(App.Path & "\file\") = "" Then
        MkDir App.Path & "\file"
    End If
    
    Open App.Path & "\file\save.txt" For Output As #1
        Print #1, "systems = {"
        Print #1, "}"
        Print #1, "planets = {"
        Print #1, "}"
    Close
    
    MsgBox "����ɹ���"
End Function

'Private Function LoadGame() As Long '������Ϸ
'    If Dir(App.Path & "\file\") = "" Then
'        MkDir App.Path & "\file"
'    End If
'
'    Open App.Path & "\file\save.txt" For Output As #1
'
'    Close
'End Function

Private Function SaveLog() As Long '�洢��Ϸ
    If Dir(App.Path & "\file\") = "" Then
        MkDir App.Path & "\file"
    End If
    
    Open App.Path & "\file\log.txt" For Output As #1
        Print #1, GameLog
    Close
End Function

'ɾ���ɴ��ϵ�����
Private Sub SpacecraftDeleteStorage(Spacecraft As Spacecraft, ID As Long)
    Dim i As Long
    
    For i = ID To UBound(Spacecraft.Storage) - 1
        Spacecraft.Storage(i) = Spacecraft.Storage(i + 1)
    Next
    ReDim Preserve Spacecraft.Storage(UBound(Spacecraft.Storage) - 1)
End Sub

'��ȡ�ɴ����ڵ�����
Private Function SpacecraftGetPlanet(Spacecraft As Spacecraft) As Long
    SpacecraftGetPlanet = SpacePositionGetPlanet(Spacecraft.Position)
End Function

'ж�طɴ��ϵ�����
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

'ж�طɴ��ϵ�����
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

'���Ʒɴ���Ӧ��UI
Private Sub SpacecraftDrawUI(Spacecraft As Spacecraft, ByVal X As Long, ByVal Y As Long)
    Dim DrawText As String
    Dim DrawRectangle As Rectangle
    Dim i As Long

    With Spacecraft
        FillColor = vbWhite
        Line (X, Y)-(X + 200, Y + 300), , B
        
        MyPrint .Name, X + 10, Y + 10
        If Not .Enabled Then MyPrint "�ѽ���", X + 10, CurrentY
        MyPrint "ά����:" & .Maintenance, X + 10, CurrentY
        MyPrint "�˿�:" & .Population, X + 10, CurrentY
        
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
        
        MyPrint "�ִ�:", X + 10, CurrentY
        
        If UBound(.Storage) = 0 Then
            MyPrint "��", X + 10, CurrentY
        Else
            For i = 1 To UBound(.Storage)
                MyPrint ResourceToString(.Storage(i)), X + 10, CurrentY
            Next
        End If
        
        DrawButtonWithUI RectangleCreate(X + 10, Y + 220, 60, 30), "���", "system(""dismantle_spacecraft " & SelectSpacecraft & """)"
        
        If .Construction = 1 Then
            '������/���ð�ť
            If .Enabled Then
                DrawText = "����"
            Else
                DrawText = "����"
            End If
            DrawButtonWithUI RectangleCreate(X + 80, Y + 220, 60, 30), DrawText, "system(""switch_spacecraft_enabled " & SelectSpacecraft & """)"
            
            DrawButtonWithUI RectangleCreate(X + 80, Y + 260, 60, 30), "ж��ȫ��", "system(""unload_spacecraft_storage " & SelectSpacecraft & """)"
            
            DrawButtonWithUI RectangleCreate(X + 80, Y + 300, 60, 30), "�򿪲ֿ�", "system(""open_spacecraft_storage " & SelectSpacecraft & """)"
            
            DrawRectangle = RectangleCreate(X + 10, Y + 260, 60, 30)
            For i = 1 To UBound(Planets)
                If Not SpacecraftGetPlanet(Spacecraft) = i Then
                    DrawButtonWithUI DrawRectangle, "�ƶ���" & Planets(i).Name, "system(""move_spacecraft_to " & SelectSpacecraft & " " & i & """)"   '�����ƶ���ť
                    DrawRectangle.Top = DrawRectangle.Top + 40
                End If
            Next
        End If
    End With
    
    DrawCloseButton X + 170, Y + 5, "clear_select"
End Sub

'������̫���е�λ��
Private Function SpacePositionCreate(PositionType As SpacePositionTypeEnum, ByVal Position1 As Double, Optional Position2 As Double = 0, Optional ByVal Progress As Double = 1) As SpacePosition
    With SpacePositionCreate
        .Type = PositionType
        .Position1 = Position1
        .Position2 = Position2
        .Progress = Progress
    End With
End Function

'��ȡ�ɴ����ڵ�����
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

'����������̫���Ĺ��
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
    Dim Freq As Currency 'ÿ���ʱ������
    Randomize
    Move Screen.Width * 0.2, Screen.Height * 0.2, Screen.Width * 0.6, Screen.Height * 0.6
    QueryPerformanceFrequency Freq
    FrequencyPerMillisecond = Freq / 1000
    FormOn = FormMainMenu
    
'    DoActionText "msgbox ������"

'    Interpreter.Run "a=4*(6+6)/-3;{b=a*2-a;c=a+b}"
    Interpreter.Run "1/0"
'    Interpreter.Run "prog a:a=3;a=7"
'    Interpreter.Run "{"
'    Interpreter.Run "system(""msgbox ������"")"
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
            '���ƴ��ڵ���Ļ
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
